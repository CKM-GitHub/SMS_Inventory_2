using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using SMS_DL;
namespace SMS_DB_Backup
{
    class Program : Base_DL
    {
        public static void Main(string[] args)
        {
            Console.Title = "DB Backup";
            back_pro("SMS_V1");        
        }

        /// <summary>
        /// Make backup Process With SQL cmd along with the created specifice dir and named the backup file name as DB names and datetime culture.
        /// </summary>
        /// <param name="name">name-string  >>> the DB name what we backup </param>
        protected static void back_pro(string name)
        {
            Base_DL bdl = new Base_DL();
            var con = bdl.GetConnectionString();
            Console.WriteLine("Test");
            var sqlConStrBuilder = new SqlConnectionStringBuilder(con);
            Console.WriteLine(sqlConStrBuilder);
            var backupFolder = SKS_pathGet("bk_path", @"C:/SMS/BackUp/");
            var io_time_on_off = " SET STATISTICS IO OFF SET STATISTICS TIME OFF SET STATISTICS IO ON SET STATISTICS TIME ON  " + " _ " + " " + Environment.NewLine + "SET STATISTICS IO OFF  SET STATISTICS TIME OFF";
            using (var connection = new SqlConnection(sqlConStrBuilder.ConnectionString))
            {
                var query = String.Format("BACKUP DATABASE {0} TO DISK='{1}'", sqlConStrBuilder.InitialCatalog, String.Format("{0}{1}-{2}.bak", backupFolder, sqlConStrBuilder.InitialCatalog, DateTime.Now.ToString("yyyy-MM-dd"))).Replace("-", "").Replace("V12", "2");
                string s = query;
                int start = s.IndexOf('\'') + 1;
                int end = s.IndexOf('\'', start);
                string result = s.Substring(start, end - start);
                if (!Directory.Exists(backupFolder))
                {
                    DirectoryInfo di = Directory.CreateDirectory(backupFolder);
                }
                if (File.Exists(result))
                {
                    File.Delete(result);
                }

                using (var command = new SqlCommand(io_time_on_off.Replace("_", query), connection))
                {
                    connection.Open();
                    command.CommandTimeout =60000;  // 180  seconds connections timeout putting make the excecution time out failed event was fixed.     // by Ptk 20190520
                    command.ExecuteNonQuery();
                }
            }
            DeleteFilesAfter7days(7, backupFolder);
        }

        /// <summary>
        /// Make Delete with custom days passed with every .bak extension 
        /// </summary>
        /// <param name="daysago"></param>
        /// <param name="bk_Path">daysago-int >> passed days , bk_path-string >> the DB name what we backup </param>
        public static void DeleteFilesAfter7days(int daysago, string bk_Path)
        {
            string[] files = Directory.GetFiles(bk_Path, "*.bak", SearchOption.TopDirectoryOnly);
            if (files.Length > 0)
            {
                string[] filesToDelete = files.Where(c =>
                {
                    TimeSpan ts = DateTime.Now - File.GetLastAccessTime(c);
                    return (ts.Days > daysago);
                }).ToArray();
                filesToDelete = filesToDelete.Count() > 0 ? filesToDelete : files.Where(c =>
                {
                    TimeSpan ts = DateTime.Now - File.GetCreationTime(c);
                    return (ts.Days > daysago);
                }).ToArray();

                foreach (string fname in filesToDelete)
                {
                    File.Delete(fname);
                }

            }
        }

    }
}
