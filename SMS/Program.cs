using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using SMS_DL;
using System.IO;
using System.Reflection;
using System.Diagnostics;


namespace SMS
{
    static class Program  
    {
        delegate int Calculator(int e);
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            //Calculator c1 = new Calculator(add);
            //c1(20);
            //var d = getNumber();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.ThreadException += Application_ThreadException;
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.Run(new Login.frmLogin());
           

        }

        
        static int number = 100;
        public static int add(int n)
        {
            number = number + n;
            return number;
        }
        public static int mul(int n)
        {
            number = number * n;
            return number;
        }
        public static int getNumber()
        {
            return number;
        }
        public static void Main1(string[] args)
        {
            Calculator c1 = new Calculator(add);//instantiating delegate  
            Calculator c2 = new Calculator(mul);
            c1(20);//calling method using delegate  
            Console.WriteLine("After c1 delegate, Number is: " + getNumber());
            c2(3);
            Console.WriteLine("After c2 delegate, Number is: " + getNumber());

        }  

        //static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        //{
        //    Base_DL bdl = new Base_DL();
        //    SqlConnection conn2 = new SqlConnection();
        //    conn2 = bdl.GetConnection();
        //    conn2.Open();

        //    SqlCommand cmd = new SqlCommand();
        //    cmd.Connection = conn2;
        ////    cmd.CommandText = "INSERT INTO Tracing_Log(CompanyName,OperatorName, PCName, Reasons, RuinedForm, [Linenumber], Path , ConflictedDateTime)   VALUES(@CompanyName,@OperatorName,@PCName, @Reasons, @RuinedForm, @Linenumber, @Path,@ConflictedDateTime)";
        //    cmd.CommandText = "INSERT INTO Tracing_Log(CompanyName,OperatorName, PCName, Reasons, RuinedForm, [Linenumber],[Path])   VALUES(@CompanyName,@OperatorName,@PCName, @Reasons, @RuinedForm, @Linenumber,@Path)";

            
            
        //    string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location.Replace("\\bin\\Debug", "")) + @"\Resources\";
        //    string[] f=  File.ReadAllText(Path.Combine(path, "store_operator")).Split('|');
            
        //    var g=(new StackTrace(e.Exception, true)).GetFrame(3).GetFileName();

        //    cmd.Parameters.AddWithValue("@CompanyName",f[0] );
        //    cmd.Parameters.AddWithValue("@OperatorName", f[1].Replace("\r\n",""));
        //    cmd.Parameters.AddWithValue("@PCName", System.Security.Principal.WindowsIdentity.GetCurrent().Name);
        //    cmd.Parameters.AddWithValue("@Reasons", e.Exception.Message);
          
        //    cmd.Parameters.AddWithValue("@RuinedForm", (new StackTrace(e.Exception, true)).GetFrame(3).GetFileName().Split('\\').Last());
        //    cmd.Parameters.AddWithValue("@Linenumber", (new StackTrace(e.Exception, true)).GetFrame(3).GetFileLineNumber());
        //    cmd.Parameters.AddWithValue("@Path",g );
        // //   cmd.Parameters.AddWithValue("@ConflictedDateTime", DateTime.Now.ToString());

        //    cmd.ExecuteNonQuery();

        //    conn2.Close();
        //}
        
    }
}
 