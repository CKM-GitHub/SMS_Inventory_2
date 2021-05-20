using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS_DL;

namespace Sugoraku_Shohinkanri_Import
{
    public partial class frmImport : Form
    {
        Base_DL bdl;
        public frmImport()
        {
            InitializeComponent();
            pb1.Minimum = 1;
            pb1.Maximum = 100;
            bdl = new Base_DL();
        }

        private void frmImport_Load(object sender, EventArgs e)
        {
            string[] strarr = { "Brand", "t_hanbai_shohin", "M_SHIIRESAKI", "Bunrui", "t_kihon_shohin", "M_HANBAI_SHOHIN", "Sport", "t_kihon_kanri", "M_KIHON_SHOHIN", "M_SOKO", "T_UKEHARAI", "T_RONRI_ZAIKO" };
          
            Task.Factory.StartNew(() =>
            {
                foreach(string str in strarr)
                {
                    Thread.Sleep(1000);
                    this.Invoke((MethodInvoker)delegate
                    {
                        switch (str)
                        {
                            case "Brand":  // 2sec
                                if (ImportProcess("1"))
                                {
                                    lblm_brand.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblm_brand.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblm_brand.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblm_brand.BackColor = Color.Red;
                                }
                                pb1.Value = 5;
                                lblNum.Text = "5";
                                lblInfo.Text = "Import Finished";
                                break;

                            case "t_hanbai_shohin":  
                                if (ImportProcess("2"))
                                {
                                    lblhanbai_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblhanbai_shohin.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblhanbai_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblhanbai_shohin.BackColor = Color.Red;
                                }
                                pb1.Value = 15;
                                lblNum.Text = "15";
                                break;

                            case "M_SHIIRESAKI":
                                if (ImportProcess("3"))
                                {
                                    lblm_shiiresaki.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblm_shiiresaki.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblm_shiiresaki.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblm_shiiresaki.BackColor = Color.Red;
                                }
                                pb1.Value = 25;
                                lblNum.Text = "25";
                                break;


                            case "Bunrui": 
                                if (ImportProcess("4"))
                                {
                                    lblm_bunrui.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblm_bunrui.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblm_bunrui.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblm_bunrui.BackColor = Color.Red;
                                }
                                pb1.Value = 30;
                                lblNum.Text = "30";
                                break;

                            case "t_kihon_shohin":
                                if (ImportProcess("5"))
                                {
                                    lblt_kihon_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblt_kihon_shohin.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblt_kihon_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblt_kihon_shohin.BackColor = Color.Red;
                                }
                                pb1.Value = 40;
                                lblNum.Text = "40";
                                break;

                            case "M_HANBAI_SHOHIN":
                                if (ImportProcess("6"))
                                {
                                    lblm_hanbai_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblm_hanbai_shohin.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblm_hanbai_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblm_hanbai_shohin.BackColor = Color.Red;
                                }
                                pb1.Value = 50;
                                lblNum.Text = "50";
                                break;

                            case "Sport":
                                if (ImportProcess("7"))
                                {
                                    lblsport.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblsport.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblsport.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblsport.BackColor = Color.Red;
                                }
                                pb1.Value = 55;
                                lblNum.Text = "55";
                                break;

                            case "t_kihon_kanri":
                                if (ImportProcess("8"))
                                {
                                    lblt_kihon_kanri.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblt_kihon_kanri.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblt_kihon_kanri.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblt_kihon_kanri.BackColor = Color.Red;
                                }
                                pb1.Value = 65;
                                lblNum.Text = "65";
                                break;

                          
                            case "M_KIHON_SHOHIN":  // 7 mins
                                if (ImportProcess("9"))
                                {
                                    lblKihon_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblKihon_shohin.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblKihon_shohin.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblKihon_shohin.BackColor = Color.Red;
                                }
                                pb1.Value = 75;
                                lblNum.Text = "75";
                                break;


                            case "M_SOKO":
                                if (ImportProcess("10"))
                                {
                                    lblm_soko.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblm_soko.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblm_soko.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblm_soko.BackColor = Color.Red;
                                }
                                pb1.Value = 80;
                                lblNum.Text = "80";
                                break;

                            case "T_UKEHARAI":
                                if (ImportProcess("11"))
                                {
                                    lblukeharai.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblukeharai.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblukeharai.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblukeharai.BackColor = Color.Red;
                                }
                                pb1.Value = 90;
                                lblNum.Text = "90";
                                break;

                            case "T_RONRI_ZAIKO":
                                if (ImportProcess("12"))
                                {
                                    lblt_ronri_zaiko.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.success;
                                    lblt_ronri_zaiko.BackColor = Color.FromArgb(192, 255, 192);
                                }
                                else
                                {
                                    lblt_ronri_zaiko.Image = Sugoraku_Shohinkanri_Import.Properties.Resources.close;
                                    lblt_ronri_zaiko.BackColor = Color.Red;
                                }
                                pb1.Value = 100;
                                lblNum.Text = "100";
                                lblInfo.Text = "Import Finished";
                                break;

                        }
                    });
                }
                Environment.Exit(0);
            });
        }

        private bool ImportProcess(string tableID)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@TableID",tableID);
            bdl.StartTransaction();
            if (bdl.InsertUpdateDeleteData(dic, "ImportProcess"))
            {
                bdl.CommitTransaction();
                return true;
            }
            return false;
        }
    }
}
