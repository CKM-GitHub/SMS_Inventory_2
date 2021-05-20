using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Data;
using SMS_Entity;
using SMS_BL;

namespace PSKS0115B_1//在庫情報反映
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.Title = "PSKS0115B_在庫情報反映";
            PSKS0115B_BL psks0115B_bl = new PSKS0115B_BL();
            DataTable dt = new DataTable();
            dt = psks0115B_bl.M_MultiPorpose_Select("108");

            while (dt.Rows[0]["Num1"].ToString().Equals("0"))
            {
                Console.WriteLine("Stop");
                Thread.Sleep(Convert.ToInt32(dt.Rows[0]["Num2"]) * 1000);
                dt = psks0115B_bl.M_MultiPorpose_Select("108");
            }

            Console.WriteLine("Start Import!");
            psks0115B_bl.PSKS0115B1_TZaiko_INSERT_UPDATE();
        }
    }
}
