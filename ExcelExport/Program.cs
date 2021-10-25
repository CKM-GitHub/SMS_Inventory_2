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
        static void Main(string[] args)
        {
            DataTable dtMail = new DataTable();
            Excel_BL excel_BL = new Excel_BL();
            dtMail = excel_BL.Mail_Select();

           
            string val1= dtMail.Rows[0]["SendedDateTime"].ToString();

            DateTime now = Convert.ToDateTime(dtMail.Rows[0]["SendedDateTime"].ToString());

            DateTime after1Month = now.AddMonths(1);
            string val2 = after1Month.ToString();
            DataTable dtExcel = excel_BL.Excel_Select(val1,val2);

            int k = 0;
            string addressKBN = string.Empty;
            if (dtExcel.Rows.Count > 0)
            {
                string FilePath = dtMail.Rows[0]["FilePath"].ToString();
                string FileFolder = dtMail.Rows[0]["FileFolder"].ToString();
                string FileName = dtMail.Rows[0]["FileName"].ToString() + ".xlsx";
                string filepath = FilePath + FileFolder+ "\\" ;
                string savefn = filepath + FileName;
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
                //if (savedialog.ShowDialog() == DialogResult.OK)
                //{
                if (Path.GetExtension(savedialog.FileName).Contains(".xlsx"))
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
                    //Process.Start(Path.GetFullPath(savedialog.FileName));
                }
                //}

            }

            if (dtMail.Rows.Count > 0)
            {
                for (int i = 0; i < dtMail.Rows.Count; i++)
                {
                    string SenderServer = "", FromMail = "", ToMail = "", CCMail = "", BCCMail = "", FromPwd = "", AttPath = "", AttFolder = "", AttFileName = "";

                    if (addressKBN != dtMail.Rows[i]["AddressKBN"].ToString())
                    {
                        addressKBN = dtMail.Rows[i]["AddressKBN"].ToString();
                        DataTable dtTemp = new DataTable();
                        dtTemp = dtMail.Select("AddressKBN='" + addressKBN + "'").CopyToDataTable();
                        MailMessage mm = new MailMessage();
                        FromMail = dtTemp.Rows[0]["SenderAddress"].ToString();
                        FromPwd = dtTemp.Rows[0]["Password"].ToString();

                        SenderServer = dtTemp.Rows[0]["SenderServer"].ToString();
                        SmtpClient smtpServer = new SmtpClient(SenderServer);
                        mm.From = new MailAddress(FromMail);

                        mm.Subject = dtTemp.Rows[0]["MailSubject"].ToString();
                        mm.Body = dtTemp.Rows[0]["MailContent"].ToString();
                        for (int ct = 0; ct < dtTemp.Rows.Count; ct++)
                        {
                            if (dtTemp.Rows[ct]["AddressKBN"].ToString().Equals("1"))
                            {
                                ToMail += dtTemp.Rows[ct]["Address"].ToString() + ",";


                            }
                            if (dtTemp.Rows[ct]["AddressKBN"].ToString().Equals("2"))
                            {
                                CCMail += dtTemp.Rows[ct]["Address"].ToString() + ",";
                                //mm.CC.Add(CCMail);

                            }
                            if (dtTemp.Rows[ct]["AddressKBN"].ToString().Equals("3"))
                            {
                                BCCMail += dtTemp.Rows[ct]["Address"].ToString() + ",";
                                //mm.Bcc.Add(BCCMail);

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
                        AttFileName = dtTemp.Rows[0]["FileName"].ToString()+ ".xlsx";

                        string filepath = AttPath + "\\" + AttFolder + "\\" + AttFileName;
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
                            if (excel_BL.MailSend_Update(Convert.ToInt32(addressKBN)))
                            {
                                Console.WriteLine("メールのご送信が完了致しました。");

                            }
                        }
                        catch (Exception ex)
                        {
                            var er = ex.Message;
                        }
                    }

                }
            }
        }
    }
}
