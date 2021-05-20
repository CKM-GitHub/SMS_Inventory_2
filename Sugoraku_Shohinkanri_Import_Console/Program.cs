using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using System.Threading;

namespace Sugoraku_Shohinkanri_Import_Console
{
    public class Program
    {
        Base_DL bdl = new Base_DL();

        public static void Main(string[] args)
        {

            string[] strarr = { "Brand", "t_hanbai_shohin", "M_SHIIRESAKI", "Bunrui", "t_kihon_shohin", "M_HANBAI_SHOHIN", "Sport", "t_kihon_kanri", "M_KIHON_SHOHIN", "M_SOKO", "T_UKEHARAI", "T_RONRI_ZAIKO", "m_nendo", "m_season" };

            foreach (string str in strarr)
            {


                Thread.Sleep(1000);


                switch (str)
                {
                    case "Brand":  // 2sec

                        if (ImportProcess("1"))
                        {
                            Console.Write("Brand import was successfully finshed in" + DateTime.Now.ToString());

                        }
                        else
                        {
                            Console.Write("Brand import was failed in" + DateTime.Now.ToString());
                        }

                        break;

                    case "t_hanbai_shohin":
                        if (ImportProcess("2"))
                        {
                            Console.Write("t_hanbai_shohin import was successfully finshed in" + DateTime.Now.ToString());

                        }
                        else
                        {
                            Console.Write("t_hanbai_shohin import was failed in" + DateTime.Now.ToString());
                        }

                        break;

                    case "M_SHIIRESAKI":
                        if (ImportProcess("3"))
                        {
                            Console.Write("M_SHIIRESAKI import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("M_SHIIRESAKI import was failed in" + DateTime.Now.ToString());
                        }

                        break;

                    case "Bunrui":
                        if (ImportProcess("4"))
                        {
                            Console.Write("Bunrui import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("Bunrui import was failed in" + DateTime.Now.ToString());
                        }

                        break;

                    case "t_kihon_shohin":
                        if (ImportProcess("5"))
                        {
                            Console.Write("t_kihon_shohin import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("t_kihon_shohin import was failed in" + DateTime.Now.ToString());
                        }

                        break;

                    case "M_HANBAI_SHOHIN":
                        if (ImportProcess("6"))
                        {
                            Console.Write("M_HANBAI_SHOHIN import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("M_HANBAI_SHOHIN import was failed in" + DateTime.Now.ToString());
                        }

                        break;


                    case "Sport":
                        if (ImportProcess("7"))
                        {
                            Console.Write("Sport import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("Sport import was failed in" + DateTime.Now.ToString());
                        }


                        break;


                    case "t_kihon_kanri":
                        if (ImportProcess("8"))
                        {
                            Console.Write("t_kihon_kanri import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("t_kihon_kanri import was failed in" + DateTime.Now.ToString());
                        }
                        break;



                    case "M_KIHON_SHOHIN":  // 7 mins
                        if (ImportProcess("9"))
                        {
                            Console.Write("M_KIHON_SHOHIN import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("M_KIHON_SHOHIN import was failed in" + DateTime.Now.ToString());
                        }

                        break;



                    case "M_SOKO":
                        if (ImportProcess("10"))
                        {
                            Console.Write("M_SOKO import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("M_SOKO import was failed in" + DateTime.Now.ToString());
                        }

                        break;


                    case "T_UKEHARAI":
                        if (ImportProcess("11"))
                        {
                            Console.Write("T_UKEHARAI import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("T_UKEHARAI import was failed in" + DateTime.Now.ToString());
                        }

                        break;


                    case "T_RONRI_ZAIKO":
                        if (ImportProcess("12"))
                        {
                            Console.Write("T_RONRI_ZAIKO import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("T_RONRI_ZAIKO import was failed in" + DateTime.Now.ToString());
                        }

                        break;
                    case "m_nendo":
                        if (ImportProcess("13"))
                        {
                            Console.Write("m_nendo import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("m_nendo import was failed in" + DateTime.Now.ToString());
                        }

                        break;

                    case "m_season":
                        if (ImportProcess("14"))
                        {
                            Console.Write("m_season import was successfully finshed in" + DateTime.Now.ToString());
                        }
                        else
                        {
                            Console.Write("m_season import was failed in" + DateTime.Now.ToString());
                        }

                        break;
                }

            }

        }
        private static bool ImportProcess(string tableID)
        {
            Base_DL bdl = new Base_DL();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@TableID", tableID);
            bdl.StartTransaction();
            if (bdl.InsertUpdateDeleteData(dic, "ImportProcess"))
            {
                ImportLog(tableID);
                bdl.CommitTransaction();
                return true;
            }
            return false;
        }

        public static void ImportLog(string tableID)
        {
            Base_DL bdl = new Base_DL();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@TableID", tableID);
            bdl.StartTransaction();
            if (bdl.InsertUpdateDeleteData(dic, "Sugoraku_Shohinkanri_Import_Log"))
            {
                bdl.CommitTransaction();

            }

        }
    }
}
