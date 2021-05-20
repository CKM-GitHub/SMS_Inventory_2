using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;
using SMS_DL;
using System.Windows;

namespace SMS_BL
{
  public class PSKS0101M_BL : Base_BL
    {
        M_MakerZaiko_H_DL hdl;
        M_MakerZaiko_D_DL ddl;
        M_NoukiHenkan_DL ndl;
        SuuryoHenKan_DL sydl;
        L_Log_DL LogDL;

        public PSKS0101M_BL()
        {
            hdl = new M_MakerZaiko_H_DL();
            ddl = new M_MakerZaiko_D_DL();
            ndl = new M_NoukiHenkan_DL();
            sydl = new SuuryoHenKan_DL();
            LogDL = new L_Log_DL();
        }

        public bool MakerZaiko_Insert(M_MakerZaiko_H_Entity makerH_data, MakerZaiko_D_Entity makerD_data, M_NoukiHenkan_Entity nouki_data, L_Log_Entity l_log, SuuryoHenKan_Entity suuryo_data)
        {
            hdl.StartTransaction();
            DataRow[] dtrowMaker = makerD_data.dtKamoKu.Select("ItemName IS NOT NULL OR ItemName <> ' '");
            DataRow[] dtrowNouki = nouki_data.dtNouki.Select("Maker IS NOT NULL OR Maker <> ' '");
            DataRow[] dtrowSuuryo = suuryo_data.dtSuuryo.Select("MakerSuuryo IS NOT NULL OR MakerSuuryo <> ' '");

            if (dtrowMaker.Count() > 0 )
            {
                DataTable dtMaker = makerD_data.dtKamoKu.Select("ItemName IS NOT NULL OR ItemName <> ' '").CopyToDataTable();
                makerD_data.KamokuXML = DataTableToXml(dtMaker);
            }
            if(dtrowNouki.Count() > 0)
            {
                DataTable dtNouki = nouki_data.dtNouki.Select("Maker IS NOT NULL OR Maker <> ' '").CopyToDataTable();
                nouki_data.NoukiXML = DataTableToXml(dtNouki);
            }

            if (dtrowSuuryo.Count() > 0)
            {
                DataTable dtSuuryo = suuryo_data.dtSuuryo.Select("MakerSuuryo IS NOT NULL OR MakerSuuryo <> ' '").CopyToDataTable();
                suuryo_data.SuuryoXML = DataTableToXml(dtSuuryo);
            }

            if (hdl.M_MakerZaiko_H_Update(makerH_data))
            {
                ddl.Transaction = hdl.Transaction;
                if (ddl.M_MakerZaiko_D_Insert(makerD_data))
                {
                    ndl.Transaction = hdl.Transaction;
                    if (ndl.M_NoukiHenkan_Insert(nouki_data))
                    {
                        sydl.Transaction = hdl.Transaction;
                        if (sydl.SuuryoHenKan_Insert(suuryo_data))
                        {
                            LogDL.Transaction = hdl.Transaction;
                            if (LogDL.L_Log_Insert(l_log))
                            {
                                hdl.CommitTransaction();
                                return true;
                            }
                        }
                    }
                }
            }
            
            return false;


        }

        public bool MakerZaiko_Update(M_MakerZaiko_H_Entity makerH_data, MakerZaiko_D_Entity makerD_data, M_NoukiHenkan_Entity nouki_data, L_Log_Entity l_log, SuuryoHenKan_Entity suuryo_data)
        {
            hdl.StartTransaction();
            DataRow[] dtrowMaker = makerD_data.dtKamoKu.Select("ItemName IS NOT NULL OR ItemName <> ' '");
            DataRow[] dtrowNouki = nouki_data.dtNouki.Select("Maker IS NOT NULL OR Maker <> ' '");
            DataRow[] dtrowSuuryo = suuryo_data.dtSuuryo.Select("MakerSuuryo IS NOT NULL OR MakerSuuryo <> ' '");

            if (dtrowMaker.Count() > 0)
            {
                DataTable dtMaker = makerD_data.dtKamoKu.Select("ItemName IS NOT NULL OR ItemName <> ' '").CopyToDataTable();
                makerD_data.KamokuXML = DataTableToXml(dtMaker);
            }
            if (dtrowNouki.Count() > 0)
            {
                DataTable dtNouki = nouki_data.dtNouki.Select("Maker IS NOT NULL OR Maker <> ' '").CopyToDataTable();
                nouki_data.NoukiXML = DataTableToXml(dtNouki);
            }

            if (dtrowSuuryo.Count() > 0)
            {
                DataTable dtSuuryo = suuryo_data.dtSuuryo.Select("MakerSuuryo IS NOT NULL OR MakerSuuryo <> ' '").CopyToDataTable();
                suuryo_data.SuuryoXML = DataTableToXml(dtSuuryo);
            }

            if (hdl.M_MakerZaiko_H_Update(makerH_data))
            {
                ddl.Transaction = hdl.Transaction;
                if (ddl.M_MakerZaiko_D_Insert(makerD_data))
                {
                    if (ndl.M_NoukiHenkan_Insert(nouki_data))
                    {
                        sydl.Transaction = hdl.Transaction;
                        if (sydl.SuuryoHenKan_Insert(suuryo_data))
                        {
                            LogDL.Transaction = hdl.Transaction;
                            if (LogDL.L_Log_Insert(l_log))
                            {
                                hdl.CommitTransaction();
                                return true;
                            }
                        }
                    }
                }

            }
          
            return false;



        }

        public bool MakerZaiko_Delete(M_MakerZaiko_H_Entity makerH_data, MakerZaiko_D_Entity makerD_data, M_NoukiHenkan_Entity nouki_data, L_Log_Entity l_log, SuuryoHenKan_Entity suuryo_data)
        {
            hdl.StartTransaction();

            makerD_data.KamokuXML = DataTableToXml(makerD_data.dtKamoKu);
            nouki_data.NoukiXML = DataTableToXml(nouki_data.dtNouki);

            if (hdl.M_MakerZaiko_H_Delete(makerH_data))
            {
                ddl.Transaction = hdl.Transaction;
                if (ddl.M_MakerZaiko_D_Delete(makerD_data))
                {
                    if (ndl.M_NoukiHenkan_Delete(nouki_data))
                    {
                        sydl.Transaction = hdl.Transaction;
                        if (sydl.SuuryoHenKan_Delete(suuryo_data))
                        {
                            LogDL.Transaction = hdl.Transaction;
                            if (LogDL.L_Log_Insert(l_log))
                            {
                                hdl.CommitTransaction();
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        public DataTable MakerZaiko_Select(MakerZaiko_D_Entity maker_D_data)
        {
            return ddl.M_MakerZaiko_D_Select(maker_D_data);
        }

        public DataTable Nouki_Select(M_NoukiHenkan_Entity nouki_data)
        {
            return ndl.M_NoukiHenkan_Select(nouki_data);
        }

        public DataTable Suuryo_Select(SuuryoHenKan_Entity suuryo_data)
        {
            return sydl.SuuryoHenKan_Select(suuryo_data);
        }
    }
}
