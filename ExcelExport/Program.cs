using ClosedXML.Excel;
using SMS_BL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelExport {
    class Program {

        static string year, month, maru = string.Empty;
        static DataTable dtMail, dtExcel = new DataTable();
        static Excel_BL excel_BL = new Excel_BL();

        public static void Main(string[] args)
        {
            dtMail = excel_BL.Mail_Select();
            string start, end ,ID= string.Empty;
            for (int i = 0; i < dtMail.Rows.Count; i++)
            {
                //ID = dtMail.Rows[i]["ID"].ToString();
                
                if (ID != dtMail.Rows[i]["ID"].ToString())
                {
                    ID = dtMail.Rows[i]["ID"].ToString();
                    start = dtMail.Rows[i]["StartDate"].ToString();
                    end = dtMail.Rows[i]["EndDate"].ToString();

                    DateTime date = Convert.ToDateTime(dtMail.Rows[i]["StartDate"]);
                    year = date.Year.ToString();
                    month = String.Format("{0:MM}", date);
                    maru = date.Month.ToString();

                    Excel(start,end);
                    if (dtExcel.Rows.Count > 0)
                    {
                        bool ret=MailSend(ID);
                        if(ret)
                            if (excel_BL.MailSend_Update(Convert.ToInt32(ID)))
                            {
                                Console.WriteLine("メールのご送信が完了致しました。");
                            }
                    }
                                       
                }

            }                  
           
        }
        private static void Excel(string start,string end)
        {
            dtExcel = excel_BL.Excel_Select(start,end);

            if (dtExcel.Rows.Count > 0)
            {
                string FilePath = dtMail.Rows[0]["FilePath"].ToString();
                string FileFolder = dtMail.Rows[0]["FileFolder"].ToString();
                string FileName = dtMail.Rows[0]["FileName"].ToString();
                string filepath = FilePath + FileFolder + "\\";
                string savefn = filepath + FileName + "（" + year + month + "）" + ".xlsx";
                if (!Directory.Exists(filepath))
                {
                    Directory.CreateDirectory(filepath);
                }
                SaveFileDialog savedialog = new SaveFileDialog();
                savedialog.Filter = "Excel Files|*.xlsx;";
                savedialog.Title = "Save";
                savedialog.FileName = FileName;
                savedialog.InitialDirectory = filepath;
                savedialog.RestoreDirectory = true;

                if (Path.GetExtension(savedialog.FileName + ".xlsx").Contains(".xlsx"))
                {
                    Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                    worksheet = workbook.ActiveSheet;
                    worksheet.Name = "worksheet";
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(dtExcel, "worksheet");
                        wb.Worksheet("worksheet").Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.SaveAs(savefn);
                    }
                }
            }
        }
        private static bool MailSend(String SenderID)
        {
            if (dtMail.Rows.Count > 0)
            {
                    string SenderServer = "", FromMail = "", ToMail = "", CCMail = "", BCCMail = "", FromPwd = "", AttPath = "", AttFolder = "", AttFileName = "";
                        DataTable dtTemp = new DataTable();
                        dtTemp = dtMail.Select("SenderID='" + SenderID + "'").CopyToDataTable();
                        MailMessage mm = new MailMessage();
                        FromMail = dtTemp.Rows[0]["SenderAddress"].ToString();
                        FromPwd = dtTemp.Rows[0]["Password"].ToString();

                        SenderServer = dtTemp.Rows[0]["SenderServer"].ToString();
                        SmtpClient smtpServer = new SmtpClient(SenderServer);
                        mm.From = new MailAddress(FromMail);

                        string s = dtTemp.Rows[0]["MailSubject"].ToString();
                        string b = dtTemp.Rows[0]["MailContent"].ToString();
                        if (s.Contains("〇") || b.Contains("〇"))
                        {
                            mm.Subject = s.Replace("〇", maru);
                            mm.Body = b.Replace("〇", maru);
                        }
                        for (int ct = 0; ct < dtTemp.Rows.Count; ct++)
                        {
                            if (dtTemp.Rows[ct]["AddressKBN"].ToString().Equals("1"))
                            {
                                ToMail += dtTemp.Rows[ct]["Address"].ToString() + ",";
                            }
                            if (dtTemp.Rows[ct]["AddressKBN"].ToString().Equals("2"))
                            {
                                CCMail += dtTemp.Rows[ct]["Address"].ToString() + ",";
                            }
                            if (dtTemp.Rows[ct]["AddressKBN"].ToString().Equals("3"))
                            {
                                BCCMail += dtTemp.Rows[ct]["Address"].ToString() + ",";
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(ToMail))
                            mm.To.Add(ToMail.TrimEnd(','));
                        if (!string.IsNullOrWhiteSpace(CCMail))
                            mm.CC.Add(CCMail.TrimEnd(','));
                        if (!string.IsNullOrWhiteSpace(BCCMail))
                            mm.Bcc.Add(BCCMail.TrimEnd(','));


                        AttPath = dtTemp.Rows[0]["FilePath"].ToString();
                        AttFolder = dtTemp.Rows[0]["FileFolder"].ToString();
                        AttFileName = dtTemp.Rows[0]["FileName"].ToString() + "（" + year + month + "）" + ".xlsx";

                        string filepath = AttPath + AttFolder + "\\" + AttFileName;
                        if (File.Exists(filepath))
                        {
                            mm.Attachments.Add(new Attachment(filepath));
                        }
                        smtpServer.Port = 587; 
                        smtpServer.Credentials = new System.Net.NetworkCredential(mm.From.Address, FromPwd);
                        smtpServer.EnableSsl = false;
                        try
                        {
                            smtpServer.Send(mm);
                            return true;                        
                        } 
                        catch (Exception ex)
                        {
                            var er = ex.Message;
                            return false;
                        }
            }
            return true;
        }
    }
}
