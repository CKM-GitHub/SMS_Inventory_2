using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using SMS_DL;
namespace ConsoleApplication1
{
    class Program :Base_DL
    {
      public  static void Main(string[] args)
      {
            back_pro("SMS_V1");
      }
       //<Summary>Make backup Process With SQL cmd along with the created specifice dir and named the backup file name as DB names and datetime culture.
        //</Summary>
        // <param > name-string  >>> the DB name what we backup </param>
      protected static void back_pro(string name)
      {
           Base_DL bdl = new Base_DL();
           var con = bdl.GetConnectionString() ;
          var sqlConStrBuilder = new SqlConnectionStringBuilder(con);
          var backupFolder = SKS_pathGet("bk_path", @"C:/SMS/BackUp/");
          using (var connection = new SqlConnection(sqlConStrBuilder.ConnectionString))
          {
              var query = String.Format("BACKUP DATABASE {0} TO DISK='{1}'", sqlConStrBuilder.InitialCatalog, String.Format("{0}{1}-{2}.bak", backupFolder, sqlConStrBuilder.InitialCatalog,DateTime.Now.ToString("yyyy-MM-dd"))).Replace("SMS_V1", name);
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

              using (var command = new SqlCommand(query, connection))
              {
                  connection.Open();
                  command.ExecuteNonQuery();
              }
          }
          DeleteFilesAfter7days(7,backupFolder);
      }
      //<Summary>Make Delete with custom days passed with every .bak extension 
      //</Summary>
      // <param > daysago-int >> passed days , bk_path-string >> Dir which include bak files  >>> the DB name what we backup </param>
      public static void DeleteFilesAfter7days(int daysago,string bk_Path)
      {
          string[] files = Directory.GetFiles(bk_Path, "*.bak", SearchOption.TopDirectoryOnly);
          if (files.Length > 0)
          {
              string[] filesToDelete = files.Where(c =>
              {
                  TimeSpan ts = DateTime.Now - File.GetLastAccessTime(c);
                  return (ts.Days > daysago);
              }).ToArray();
              foreach(string fname in filesToDelete )
              {
                  File.Delete(fname);
              }
          }
      }

    }
}
