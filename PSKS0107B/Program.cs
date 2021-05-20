using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_BL;
using SMS_Entity;
using System.Threading;
using System.IO;
using System.Collections;


namespace PSKS0107B
{
    class Program
    {
       public  static void Main(string[] args)
        {
            Console.Title = "PSKS0107B_入出庫情報取得";
            Base_Entity baseEntity = new Base_Entity();
            PSKS0107B_BL psks0107b_bl = new PSKS0107B_BL();
            M_MultiPorpose_Entity mmultiporpose_data = new M_MultiPorpose_Entity();
            DataTable dt = new DataTable();

            mmultiporpose_data.ID = "103";
            mmultiporpose_data.Key = "1";
            dt = psks0107b_bl.M_MultiPorpose_SelectID(mmultiporpose_data);

            while (dt.Rows[0]["Num1"].ToString().Equals("0"))
            {
                Thread.Sleep(Convert.ToInt32(dt.Rows[0]["Num2"])*1000);
                dt = psks0107b_bl.M_MultiPorpose_SelectID(mmultiporpose_data); // select columns from M_Multiporpose table
            }
           
            //DataTable dtpostgre = new DataTable();
            //dtpostgre = psks0107b_bl.t_hanbai_shohin_Select();

            //DataTable dtsukoraku = new DataTable();
            //dtsukoraku = psks0107b_bl.Sukoraku_Select();

            //DataTable dtukeharai = new DataTable();

            //dtukeharai = psks0107b_bl.UKeHaRai_Select();

            bool result = false;
            string datetime = DateTime.Now.ToString();
            baseEntity = new Base_Entity();
            baseEntity.InsertDateTime = datetime;
            result = psks0107b_bl.Data_Insert(baseEntity);

        }
    }
}
