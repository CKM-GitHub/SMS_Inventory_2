using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SMS_BL;
using SMS_Entity;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using System.Configuration;

namespace PSKS0113B
{
    class Program
    {
        //static String sourcePath = ConfigurationManager.AppSettings["SourcePath"].ToString();
        //static String targetPath = ConfigurationManager.AppSettings["TargetPath"].ToString();

        public static void Main(string[] args)
        {
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
            System.IO.File.Copy(s, destFile, true);

            #endregion

            Console.WriteLine("Import Start!");
            if (PSKS0113Bbl.ImportShopData())
            {
                System.IO.File.Delete(destFile); // delete file after import data
                Console.WriteLine("Finished");
            }
            else
            {
                Console.WriteLine("Error");
                Console.ReadLine();
            }
        }
    }
}
