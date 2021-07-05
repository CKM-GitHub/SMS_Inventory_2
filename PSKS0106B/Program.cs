using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;
using SMS_BL;
using System.Threading;
using System.IO;
using System.Collections;
using System.Net;


namespace PSKS0106B
{
    class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                Console.Title = "PSKS0106B_販売情報取得";
                PSKS0106B_BL psks_bl = new PSKS0106B_BL();
                DataTable dt = new DataTable();
                M_MultiPorpose_Entity mmultiporpose_data = new M_MultiPorpose_Entity();
                mmultiporpose_data.ID = "102";
                mmultiporpose_data.Key = "1";
                dt = psks_bl.M_MultiPorpose_SelectID(mmultiporpose_data);

                while (dt.Rows[0]["Num1"].ToString().Equals("0"))
                {
                    Thread.Sleep(Convert.ToInt32(dt.Rows[0]["Num2"]) * 1000);
                    dt = psks_bl.M_MultiPorpose_SelectID(mmultiporpose_data); // select columns from M_Multiporpose table
                }

                #region CSV Download from SKS FTP

                string[] files = GetFileList();
                if (files != null)
                {
                    foreach (string file in files)
                    {
                        Download(file);
                    }
                }
                #endregion

                #region Import to T_Ukebarai

                ArrayList arrlst = CheckCsvHeader(dt.Rows[0]["Char1"].ToString()); // check column_headers in csv file

                DataTable dtCsvAll = new DataTable();
                foreach (string data in arrlst)  // read csv file
                {

                    string path = dt.Rows[0]["Char1"].ToString() + @"\" + data;


                    string[] Lines = File.ReadAllLines(path, Encoding.GetEncoding(932));
                    string[] Fields;
                    Fields = Lines[0].Split(new char[] { ',' });
                    int Cols = Fields.GetLength(0);
                    DataTable dtcsv = new DataTable();
                    //1st row must be column names; force lower case to ensure matching later on.
                    for (int i = 0; i < Cols; i++)
                        dtcsv.Columns.Add(Fields[i].Replace("\"", ""), typeof(string));
                    DataRow Row;
                    for (int i = 1; i < Lines.GetLength(0); i++)
                    {
                        Fields = Lines[i].Split(new char[] { ',' });
                        Row = dtcsv.NewRow();
                        for (int f = 0; f < Cols; f++)
                            Row[f] = Fields[f].Replace("\"", "");
                        dtcsv.Rows.Add(Row);
                    }

                    dtCsvAll.Merge(dtcsv);
                }

                if (dtCsvAll.Rows.Count > 0)
                {
                    bool result = false;
                    result = psks_bl.Ukebarai_Insert(dtCsvAll); // insert into ukebarai

                }
                #endregion

                #region Move to backup foler finished Files

                for (int i = 0; i < arrlst.Count; i++)  // move csv file to another folder
                {
                    string path = dt.Rows[0]["Char1"].ToString() + @"\" + arrlst[i];
                    string movefile = dt.Rows[0]["Char1"].ToString() + @"\" + "Finish";
                    //string movefile = dt.Rows[0]["Char2"].ToString();
                    if (!Directory.Exists(@movefile))
                        Directory.CreateDirectory(@movefile);

                    string destinationFile = movefile + @"\" + arrlst[i];

                    if (System.IO.File.Exists(destinationFile))
                    {
                        System.IO.File.Delete(destinationFile);
                    }

                    System.IO.File.Move(path, destinationFile);
                }
                #endregion

            }
            catch (Exception e)
            {
                InsertErrorLog(e.Message.ToString());
            }
        }

        public static string[] GetFileList()
        {
            string[] downloadFiles;
            StringBuilder result = new StringBuilder();
            WebResponse response = null;
            StreamReader reader = null;
            try
            {
                FtpWebRequest reqFTP;
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://163.43.113.92:22/"));
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential("capital_order_mail", "c!PiTal!0rd@rMa!l");
                reqFTP.Method = WebRequestMethods.Ftp.ListDirectory;
                reqFTP.Proxy = null;
                reqFTP.KeepAlive = false;
                //reqFTP.UsePassive = false;
                response = reqFTP.GetResponse();
                reader = new StreamReader(response.GetResponseStream());
                string line = reader.ReadLine();
                while (line != null)
                {
                    result.Append(line);
                    result.Append("\n");
                    line = reader.ReadLine();
                }
                // to remove the trailing '\n'
                result.Remove(result.ToString().LastIndexOf('\n'), 1);
                return result.ToString().Split('\n');
            }
            catch
            {
                if (reader != null)
                {
                    reader.Close();
                }
                if (response != null)
                {
                    response.Close();
                }
                downloadFiles = null;
                return downloadFiles;


            }
        }

        private static void Download(string file)
        {
            try
            {
                PSKS0106B_BL psks_bl = new PSKS0106B_BL();

                //string uri = "ftp://" + "153.120.180.197" + "/" + "Inventory_FromSKS" + "/" + file;
                string uri = "ftp://" + "163.43.113.92:22" + "/" + file;
                Uri serverUri = new Uri(uri);
                if (serverUri.Scheme != Uri.UriSchemeFtp)
                {
                    return;
                }
                FtpWebRequest reqFTP;
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://" + "163.43.113.92:22" + "/" + file));
                reqFTP.Credentials = new NetworkCredential("capital_order_mail", "c!PiTal!0rd@rMa!l");
                reqFTP.KeepAlive = false;
                reqFTP.Method = WebRequestMethods.Ftp.DownloadFile;
                reqFTP.UseBinary = true;
                reqFTP.Proxy = null;
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream responseStream = response.GetResponseStream();

                DataTable dt = new DataTable();

                dt = psks_bl.M_MultiPorpose_SelectChar();

                string path = dt.Rows[0]["Char1"].ToString();
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                FileStream writeStream = new FileStream(dt.Rows[0]["Char1"].ToString() + "/" + file, FileMode.Create);
                int Length = 2048;
                Byte[] buffer = new Byte[Length];
                int bytesRead = responseStream.Read(buffer, 0, Length);
                while (bytesRead > 0)
                {
                    writeStream.Write(buffer, 0, bytesRead);
                    bytesRead = responseStream.Read(buffer, 0, Length);
                }

                writeStream.Close();
                response.Close();
                FtpWebRequest requestFileDelete = (FtpWebRequest)WebRequest.Create("ftp://163.43.113.92:22/" + file);
                requestFileDelete.Credentials = new NetworkCredential("capital_order_mail", "c!PiTal!0rd@rMa!l");
                requestFileDelete.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse responseFileDelete = (FtpWebResponse)requestFileDelete.GetResponse();
            }

            catch
            {
                //MessageBox.Show(ex.Message, "Download Error");
            }

        }

        static void InsertErrorLog(string errormsg)
        {
            Message_BL mes_bl = new Message_BL();
            mes_bl.INSERT_ErrorTable(errormsg, "PSKS0106B");
        }


        public static ArrayList CheckCsvHeader(string path)//left file that has wrong column header 
        {
            ArrayList arrlst = new ArrayList();

            string[] csvfiles = Directory.GetFiles(path, "*.csv")
                                      .Select(Path.GetFileName)
                                      .ToArray();

            foreach (string str in csvfiles)
            {
                string filepath = path + @"\" + str;
                string[] lines = File.ReadAllLines(filepath, Encoding.GetEncoding(932));
                string[] fields = lines[0].Split(',');
                if (fields.Contains("\"WEBストア名\"") && fields.Contains("\"販売日時\"") && fields.Contains("\"商品CD\"") && fields.Contains("\"数量\"") && fields.Contains("\"受注番号\"") && fields.Contains("\"AdminCD\""))
                    arrlst.Add(str);
            }

            return arrlst;

        }

    }
}