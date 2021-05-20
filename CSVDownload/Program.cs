using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections ;
using System.Data;
using System.Net;
using SMS_BL;
using SMS_Entity;

namespace CSVDownload
{
    class Program
    {
        public static void Main(string[] args)
        {

            string[] files = GetFileList();
            if (files != null )
            {
                foreach (string file in files)
                {
                    Download(file);
                }
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
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://153.120.180.197/Inventory_FromSKS/"));
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential("master_import", "d08i34wi@dd34");
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
               
                string uri = "ftp://" + "153.120.180.197" + "/" + "Inventory_FromSKS" + "/" + file;
                Uri serverUri = new Uri(uri);
                if (serverUri.Scheme != Uri.UriSchemeFtp)
                {
                    return;
                }       
                FtpWebRequest reqFTP;                
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://" + "153.120.180.197" + "/" + "Inventory_FromSKS"+ "/" + file));                                
                reqFTP.Credentials = new NetworkCredential("master_import", "d08i34wi@dd34");                
                reqFTP.KeepAlive = false;                
                reqFTP.Method = WebRequestMethods.Ftp.DownloadFile;                                
                reqFTP.UseBinary = true;
                reqFTP.Proxy = null;                 
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream responseStream = response.GetResponseStream();

                DataTable dt = new DataTable();
               
                dt = psks_bl.M_MultiPorpose_SelectChar();

                 string path = dt.Rows [0]["Char1"].ToString ();
                 if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                FileStream writeStream = new FileStream(dt.Rows[0]["Char1"].ToString() +"/" + file, FileMode.Create);                
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
                FtpWebRequest requestFileDelete = (FtpWebRequest)WebRequest.Create("ftp://153.120.180.197/Inventory_FromSKS/" + file);
                requestFileDelete.Credentials = new NetworkCredential("master_import", "d08i34wi@dd34");
                requestFileDelete.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse responseFileDelete = (FtpWebResponse)requestFileDelete.GetResponse();
            }
           
            catch
            {
                //MessageBox.Show(ex.Message, "Download Error");
            }

        }
  
    }
}