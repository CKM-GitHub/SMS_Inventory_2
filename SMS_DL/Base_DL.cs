using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;

namespace SMS_DL
{
    public class Base_DL
    {
        #region Variables
        public SqlConnection connection;
        public static SqlConnection connection2;
        protected SqlCommand command;
        protected SqlTransaction transaction;
        protected SqlDataAdapter adapter;
        protected bool useTransaction;
        
        #endregion

        #region Properties
        public SqlTransaction Transaction
        {
            get { return transaction; }
            set { transaction = value; }
        }

        public bool UseTransaction
        {
            get { return useTransaction; }
            set { useTransaction = value; }
        }
        #endregion

        #region Constructors
        public Base_DL()
        {
            //this.connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBMain"].ConnectionString);
        }
        #endregion


        public static string SKS_Path()
        {
            return SKS_pathGet("Sks_Path", "http://www.capital-sys.net/RCM_Capital/WebForms/Import/Json_Import_Test.ashx?User=system&Password=capital");
        }
        public void ChnageDB_2014(bool flg)
        {
            if (flg)
            { 
                
            }

        }
        public  string path_108i(string Key, string path)
        {
            return SKS_pathGet(Key,path);
        }
        public SqlConnection GetConnection()
        {
            if (connection == null)
                connection = new SqlConnection(GetConnectionString());

            return connection;
        }
        #region Transaction Methods
        public void StartTransaction()
        {
            connection = GetConnection();
            if (connection.State == ConnectionState.Closed)
                connection.Open();

            transaction = connection.BeginTransaction();
            useTransaction = true;
        }

        public void CommitTransaction()
        {
            connection = GetConnection();
            if (transaction != null)
                transaction.Commit();

            if (connection.State == ConnectionState.Open)
                connection.Close();

            useTransaction = false;
        }

        public void RollBackTransaction()
        {
            connection = GetConnection();
            if (transaction != null)
                transaction.Rollback();

            if (connection.State == ConnectionState.Open)
                connection.Close();

            useTransaction = false;
        }
        #endregion

        protected void AddParam(SqlCommand cmd, string key, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                cmd.Parameters.AddWithValue(key, DBNull.Value);
            else
                cmd.Parameters.AddWithValue(key, value);
        }

        public DataTable DynamicSelectData(string fields, string tableName, string condition)
        {
            string query = "Select " + fields + " From " + tableName + " " + condition;
            SqlConnection conn = GetConnection();
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            da.Fill(dt);
            conn.Close();
            return dt;
        }

        public bool DynamicUpdateData(M_MultiPorpose_Entity mme)
        {
            string query =  "Update " + mme.tableName + " Set Char1= '" + mme.fields.Split(',').First()  + "' , Char2= '" + mme.fields .Split (',').Last () +"'" + mme.condition ;
            SqlConnection conn = GetConnection();
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.CommandTimeout = 0;
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                if (!useTransaction)
                {
                    cmd.Connection.Close();
                }
                return true;
            }
            catch
            {
                return false;
            }


            
        }

        public DataTable SelectData(Dictionary<string, string> dic, string sp)
        
        {
            DataTable dt = new DataTable();
            if (useTransaction)
                command = new SqlCommand(sp, GetConnection(), transaction);
            else
                command = new SqlCommand(sp, GetConnection());
            command.CommandType = CommandType.StoredProcedure;
            command.CommandTimeout = 0;
            foreach (KeyValuePair<string, string> pair in dic)
            {
                AddParam(command, pair.Key, (pair.Value == null) ? pair.Value : pair.Value.Trim());
            }
            adapter = new SqlDataAdapter(command);
            adapter.Fill(dt);

            return dt;
        
        }
        public DataSet SelectDS(Dictionary<string, string> dic, string sp)
        {
            DataSet dt = new DataSet();
            if (useTransaction)
                command = new SqlCommand(sp, GetConnection(), transaction);
            else
                command = new SqlCommand(sp, GetConnection());
            command.CommandType = CommandType.StoredProcedure;
            command.CommandTimeout = 0;
            foreach (KeyValuePair<string, string> pair in dic)
            {
                AddParam(command, pair.Key, (pair.Value == null) ? pair.Value : pair.Value.Trim());
            }
            adapter = new SqlDataAdapter(command);
            adapter.Fill(dt);

            return dt;

        }
        public DataSet SelectDataSet(Dictionary<string, string> dic, string sp)
        {
            DataSet dt = new DataSet();
            if (useTransaction)
                command = new SqlCommand(sp, GetConnection(), transaction);
            else
                command = new SqlCommand(sp, GetConnection());
            command.CommandType = CommandType.StoredProcedure;
            command.CommandTimeout = 0;
            foreach (KeyValuePair<string, string> pair in dic)
            {
                AddParam(command, pair.Key, (pair.Value == null) ? pair.Value : pair.Value.Trim());
            }
            adapter = new SqlDataAdapter(command);
            adapter.Fill(dt);

            return dt;

        }
        public bool InsertUpdateDeleteData(Dictionary<string, string> dic, string sp)
        {
            try
            {
                if (useTransaction)
                    command = new SqlCommand(sp, GetConnection(), transaction);
                else
                    command = new SqlCommand(sp, GetConnection());
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 0;
                foreach (KeyValuePair<string, string> pair in dic)
                {
                    AddParam(command, pair.Key, (pair.Value == null) ? pair.Value : pair.Value.Trim());
                }
             
                if (!useTransaction)
                    command.Connection.Open();
                command.ExecuteNonQuery();
                if (!useTransaction)
                    command.Connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                RollBackTransaction();
                return false;
            }
        }

        public int InsertData(Dictionary<string, string> dic, string sp)
        {
            try
            {
                if (useTransaction)
                    command = new SqlCommand(sp, GetConnection(), transaction);
                else
                    command = new SqlCommand(sp, GetConnection());
                command.CommandType = CommandType.StoredProcedure;
                foreach (KeyValuePair<string, string> pair in dic)
                {
                    AddParam(command, pair.Key, (pair.Value == null) ? pair.Value : pair.Value.Trim());
                }

                if (!useTransaction)
                    command.Connection.Open();
                int id = (int)command.ExecuteScalar();
                if (!useTransaction)
                    command.Connection.Close();
                return id;
            }
            catch (Exception)
            {
                return -1;
            }
        }

        public bool CheckExists(Dictionary<string, string> dic, string sp)
        {
            DataTable dt = SelectData(dic, sp);
            return dt.Rows.Count > 0 ? true : false;
        }

        public DataTable SelectData_CommandText(string commandText)
        {
            DataTable dt = new DataTable();
            if (useTransaction)
                command = new SqlCommand(commandText, GetConnection(), transaction);
            else
                command = new SqlCommand(commandText, GetConnection());
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            adapter = new SqlDataAdapter(command);
            adapter.Fill(dt);

            return dt;
        }

        public bool INSERT_ErrorTable(string error, string formName)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@ErrorMsg", error);
            dic.Add("@FormName", formName);
            return InsertUpdateDeleteData(dic, "INSERT_ErrorTable");
        }

        public void Change_DBVersion()
        {
            
            if (!Directory.Exists(@"C:\SMS\Config"))
                Directory.CreateDirectory(@"C:\SMS\Config");
            if (!File.Exists(@"C:\SMS\Config\App.config"))
                File.Create(@"C:\SMS\Config\App.config");

            Configuration config = ConfigurationManager.OpenExeConfiguration(@"C:\SMS\Config\App.config");
       //     ConnectionStringSettings setting=null;
            //if (config.ConnectionStrings.ConnectionStrings["SMS_DB"] != null)
            //{
            //  //setting = new ConnectionStringSettings("SMS_DB", "Data Source=WIN-0LHAOSODL41\\SQLEXPRESS;Initial Catalog=SMS_V1;Persist Security Info=True;User ID=sa;Password=admin123456!", "System.Data.SqlClient");
            //   config.ConnectionStrings.ConnectionStrings["SMS_DB"].ConnectionString = "Data Source=WIN-0LHAOSODL41\\SQL2014;Initial Catalog=SMS_V1;Persist Security Info=True;User ID=sa;Password=admin123456!";   // For server
            //    //config.ConnectionStrings.ConnectionStrings["SMS_DB"].ConnectionString = "Data Source=DEVSERVER\\SQLEXPRESS;Initial Catalog=SMS_V1;Persist Security Info=True;User ID=sa;Password=admin123456!";
              
            //    config.Save(ConfigurationSaveMode.Modified);
            //   ConfigurationManager.RefreshSection("connectionStrings");
            //}

        }

        public string GetConnectionString()
        {
            if (!Directory.Exists(@"C:\SMS\Config"))
                Directory.CreateDirectory(@"C:\SMS\Config");
            if (!File.Exists(@"C:\SMS\Config\App.config"))
                File.Create(@"C:\SMS\Config\App.config");

            Configuration config = ConfigurationManager.OpenExeConfiguration(@"C:\SMS\Config\App.config");
           // config.ConnectionStrings.ConnectionStrings["SMS_DB"] is for origianl
            if (config.ConnectionStrings.ConnectionStrings["SMS_DB"] == null)           // config.ConnectionStrings.ConnectionStrings["SMS_DB"] is for origianl
            {
                ConnectionStringSettings setting = new ConnectionStringSettings("SMS_DB", "Data Source=WIN-0LHAOSODL41\\SQL2014;Initial Catalog=SalesManagement;Persist Security Info=True;User ID=sa;Password=admin123456!", "System.Data.SqlClient");
                config.ConnectionStrings.ConnectionStrings.Add(setting);
                config.Save(ConfigurationSaveMode.Modified);
            }

            string constr = config.ConnectionStrings.ConnectionStrings["SMS_DB"].ToString();
            return constr;
        }
        public static string GetConnectionString2()
        {
            if (!Directory.Exists(@"C:\SMS\Config"))
                Directory.CreateDirectory(@"C:\SMS\Config");
            if (!File.Exists(@"C:\SMS\Config\App.config"))
                File.Create(@"C:\SMS\Config\App.config");

            Configuration config = ConfigurationManager.OpenExeConfiguration(@"C:\SMS\Config\App.config");

            if (config.ConnectionStrings.ConnectionStrings["SMS_DB"] == null)
            {
                ConnectionStringSettings setting = new ConnectionStringSettings("SMS_DB", "Data Source=DEVSERVER\\SQLEXPRESS;Initial Catalog=SalesManagement;Persist Security Info=True;User ID=sa;Password=12345", "System.Data.SqlClient");
                config.ConnectionStrings.ConnectionStrings.Add(setting);
                config.Save(ConfigurationSaveMode.Modified);
            }

            string constr = config.ConnectionStrings.ConnectionStrings["SMS_DB"].ToString();
            return constr;
        }

        public static string SKS_pathGet(string key, string value)
        {

            Configuration config = ConfigurationManager.OpenExeConfiguration(@"C:\SMS\Config\App.config");

            if (config.AppSettings.Settings[key] == null)
            {
               
            config.AppSettings.Settings.Add(key, value);
            config.Save(ConfigurationSaveMode.Modified);
           
            }
            return config.AppSettings.Settings[key].Value;

        }
    }
}
