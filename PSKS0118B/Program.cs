using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

using System.IO;using System.Collections;
using System.Data;
using SMS_BL;

namespace PSKS0118B
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.Title = "PSKS0118B_在庫FLG確認";
            PSKS0118B_BL psks0118B_bl = new PSKS0118B_BL();
            DataTable dt = new DataTable();
            dt = psks0118B_bl.M_MultiPorpose_Select("109");

            while (dt.Rows[0]["Num1"].ToString().Equals("0"))
            {
                Console.WriteLine("Stop");
                Thread.Sleep(Convert.ToInt32(dt.Rows[0]["Num2"]) * 1000);
                dt = psks0118B_bl.M_MultiPorpose_Select("109");
            }

            PSKS0118B_BL psks0118b_bl = new PSKS0118B_BL();
            bool result = false;
            result = psks0118b_bl.SKSChangeFLG();
        }
    }
}
