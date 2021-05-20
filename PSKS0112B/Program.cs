
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_BL;
using SMS_Entity;
using System.Data;
using System.Threading;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using SMS_DL;
using System.Configuration;




namespace PSKS0112B
{
    class Program : Base_DL
    {

        public static void Main(string[] args)
        {
            Console.Title = "PSKS0112B_SKS在庫連携";
            M_MultiPorpose_Entity mmpData = new M_MultiPorpose_Entity();
            PSKS0112B_BL psks0112bl = new PSKS0112B_BL();
            DataTable dt = new DataTable();

            string datetime = string.Empty;

            dt = psks0112bl.M_MultiPorpose_Select();

            while (dt.Rows[0]["Num1"].ToString().Equals("0"))
            {
                Thread.Sleep(Convert.ToInt32(dt.Rows[0]["Num2"]) * 1000);
                dt = psks0112bl.M_MultiPorpose_Select();
            }
            Base_Entity be = new Base_Entity();
            be.InsertDateTime = DateTime.Now.ToString();
            dt = psks0112bl.PSKS0112B_Select(be);
            string JS_result = datatableToJason(dt);
            PostJS(JS_result);
        }

        public static string datatableToJason(DataTable dt)
        {
            string JS = string.Empty;
            JS = "{\"inventorynew\":"+JsonConvert.SerializeObject(dt)+"}";
            return JS;
        }

        public static void PostJS(string re)
        {
            Base_DL bdl = new Base_DL();
            var baseAddress = SKS_Path();
            var http = (HttpWebRequest)WebRequest.Create(new Uri(baseAddress));
            http.Accept = "application/json";
            http.ContentType = "application/json";
            http.Method = "POST";
            string parsedContent = re;
            Encoding encoding = Encoding.UTF8;
            //ASCIIEncoding encoding = new ASCIIEncoding();
            Byte[] bytes = encoding.GetBytes(parsedContent);
            http.Timeout = System.Threading.Timeout.Infinite;

            http.ReadWriteTimeout = System.Threading.Timeout.Infinite;
            Stream newStream = http.GetRequestStream();
            newStream.Write(bytes, 0, bytes.Length);
            newStream.Close();

            var response = http.GetResponse();
            var stream = response.GetResponseStream();
            var sr = new StreamReader(stream);
            var content = sr.ReadToEnd();
        }

    
       
    }
}
