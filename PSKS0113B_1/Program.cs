using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_BL;
using SMS_Entity;
using System.Threading;
using System.IO;

namespace PSKS0113B_1
{
    class Program
    {

        static void Main(string[] args)
        {


            Console.Title = "PSKS0113B_店舗在庫情報取得";
            String sourcePath = @"\\192.168.0.14\ablesoft\CONPHAS-GENERAL\DATA"; 
            String targetPath = @"F:\Development\ShopData";
            M_MultiPorpose_Entity mme = new M_MultiPorpose_Entity();
            PSKS0113B_BL PSKS0113Bbl = new PSKS0113B_BL();

            mme = PSKS0113Bbl.M_MultiPurpose_Select();

            while (mme.Num1 == "0")
            {
                Console.WriteLine("Stop"); 
                Thread.Sleep(Convert.ToInt32(mme.Num2) * 1000);
                mme = PSKS0113Bbl.M_MultiPurpose_Select();
            }

            #region copy file

            string fileName = "FreeDAT.mdb";
            string s = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);
            string shopFlag = "0";
            if (File.Exists(s))
            {
                System.IO.File.Copy(s, destFile, true);
                shopFlag = "1";
            }

            #endregion

            Console.WriteLine("Import Start!");
            if (PSKS0113Bbl.ImportShopData(shopFlag))
            {
                if(shopFlag.Equals("1"))
                    System.IO.File.Delete(destFile); // delete file after import data
            }
            else
            {
            }
        }
    }
}
