using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using System.Collections;

namespace SMS_BL
{
    public class PSKS0103I_BL : Base_BL
    {
        M_MultiPorpose_DL mmpdl;
        T_MakerInportRireki_DL tmirdl;
        T_MakerInport_DL tmidl;
        PSKS0103I_DL psks0103idl;
        M_MakerZaiko_D_DL mmzddl;
        M_SiiresakiZaiko_DL mszdl;
        Base_DL bdl;
        T_MakerInport_Entity tmie;
        public DataTable get_ptnname(string mkr_CD)
        {
        return mmpdl.get_ptnname(mkr_CD);
        }
        public PSKS0103I_BL()
        {
            bdl = new Base_DL();
            mmpdl = new M_MultiPorpose_DL();
            tmirdl = new T_MakerInportRireki_DL();
            psks0103idl = new PSKS0103I_DL();
            tmidl = new T_MakerInport_DL();
            mmzddl = new M_MakerZaiko_D_DL();
            mszdl = new M_SiiresakiZaiko_DL();
        }

        public DataTable M_MultiPorpose_DynamicSelect(M_MultiPorpose_Entity mme)
        {
            DataTable dt = mmpdl.DynamicSelectData(mme.fields, mme.tableName, mme.condition);
            return dt;
        }

        public bool M_MultiPorpose_DynamicUpdateData(M_MultiPorpose_Entity mme)
        {
            return mmpdl.DynamicUpdateData(mme);
        }

        public DataTable T_MakerInportRireki_SELECT(T_MakerInportRireki_Entity tmire)
        {
            return tmirdl.T_MakerInportRireki_SELECT(tmire);
        }

        public bool ImportProcess(T_MakerZaiko_Entity tmze, string patternCD, DataTable dtPattern, int i, string insertDateTime)
        {
            psks0103idl.StartTransaction();
            bool result = false;
            if (i == 0)
            {
                result = Delete_BeforeImoprt(tmze);//do for first file only
            }
            else
                result = true;

            if (result)
            {
                if (InportProcesses(tmze,patternCD,dtPattern, insertDateTime))
                {
                    psks0103idl.CommitTransaction();
                    return true;
                }
            }
            return false;
        }

        public bool Delete_BeforeImoprt(T_MakerZaiko_Entity tmze)//取込前データ削除
        {
            M_SiiresakiZaiko_DL mszdl = new M_SiiresakiZaiko_DL();
            M_SiiresakiZaiko_Entity msze = new M_SiiresakiZaiko_Entity();
            T_MakerZaiko_DL tmzdl = new T_MakerZaiko_DL();
            msze.SiiresakiCD = tmze.MakerCD;
            DataTable dt = mszdl.M_SiiresakiZaiko_Select(msze);//確定済の判断
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["TorikomiFixedFLG"].ToString() == "1")//TorikomiFixedFLG == 1なら(確定済)
                {
                    tmzdl.Transaction = psks0103idl.Transaction;
                    if (tmzdl.T_MakerZaiko_DELETE(tmze))//既存データ削除
                    {
                        tmze.InsertDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                        if (!tmzdl.T_MakerZaiko_Insert(tmze))
                            return false;
                    }
                    else return false;
                }
            }
            return true;
        }

        public DataTable RemoveDuplicateRows(DataTable dTable, string colName)
        {
            Hashtable hTable = new Hashtable();
            ArrayList duplicateList = new ArrayList();

            //Add list of all the unique item value to hashtable, which stores combination of key, value pair.
            //And add duplicate item value in arraylist.
            foreach (DataRow drow in dTable.Rows)
            {
                if (hTable.Contains(drow[colName]))
                    duplicateList.Add(drow);
                else
                    hTable.Add(drow[colName], string.Empty);
            }

            //Removing a list of duplicate items from datatable.
            foreach (DataRow dRow in duplicateList)
                dTable.Rows.Remove(dRow);

            //Datatable which contains unique records will be return as output.
            return dTable;
        }

        private bool InportProcesses(T_MakerZaiko_Entity tmze, string patternCD, DataTable dtPattern, string insertDateTime)//取込処理
        {
            try
            {
                string noukiFlag = "0";

                DataRow[] drs = dtPattern.Select("Convert(ItemProperty, 'System.Int32') = 1");
                int colIndex = Convert.ToInt32(drs[0]["SEQ"].ToString()) - 1;

                tmze.dt1.Columns[colIndex].ColumnName = "JANCD";

                drs = dtPattern.Select("Convert(ItemProperty, 'System.Int32') = 2");
                int i = 0;
                foreach (DataRow dr in drs)
                {
                    colIndex = Convert.ToInt32(drs[i]["SEQ"].ToString()) - 1;
                    tmze.dt1.Columns[colIndex].ColumnName = "P" + ++i;
                }
                while (i < 5)
                {
                    DataColumn dc = new DataColumn();
                    dc.DefaultValue = 0;
                    dc.ColumnName = "P" + ++i;
                    tmze.dt1.Columns.Add(dc);
                }

                i = 0;
                drs = dtPattern.Select("Convert(ItemProperty, 'System.Int32') = 3");
                foreach (DataRow dr in drs)
                {
                    colIndex = Convert.ToInt32(drs[i]["SEQ"].ToString()) - 1;
                    tmze.dt1.Columns[colIndex].ColumnName = "M" + ++i;
                }

                while (i < 5)
                {
                    DataColumn dc = new DataColumn();
                    dc.DefaultValue = 0;
                    dc.ColumnName = "M" + ++i;
                    tmze.dt1.Columns.Add(dc);
                }

                tmze.dt1.Columns.Add("ZaikoSuTotal");

                drs = dtPattern.Select("Convert(ItemProperty, 'System.Int32') = 9 or Convert(ItemProperty, 'System.Int32') = 10");
                if (drs.Count() > 0)
                {
                    colIndex = Convert.ToInt32(drs[0]["SEQ"].ToString()) - 1;
                    tmze.dt1.Columns[colIndex].ColumnName = "Nouki";

                    if (drs[0]["ItemProperty"].ToString().Equals("9"))
                        noukiFlag = "1";
                }
                else
                    tmze.dt1.Columns.Add("Nouki");

                ArrayList arrcol = new ArrayList();

                foreach (DataColumn dc in tmze.dt1.Columns)
                {
                    if (!(dc.ColumnName.Equals("JANCD") || dc.ColumnName.Equals("P1") || dc.ColumnName.Equals("P2") || dc.ColumnName.Equals("P3") || dc.ColumnName.Equals("P4") || dc.ColumnName.Equals("P5") ||
                       dc.ColumnName.Equals("M1") || dc.ColumnName.Equals("M2") || dc.ColumnName.Equals("M3") || dc.ColumnName.Equals("M4") || dc.ColumnName.Equals("M5") || dc.ColumnName.Equals("ZaikoSuTotal") ||
                       dc.ColumnName.Equals("Nouki")))
                    {
                        arrcol.Add(dc);
                    }
                }

                foreach (DataColumn dc in arrcol)
                    tmze.dt1.Columns.Remove(dc);

                tmze.dt1.AcceptChanges();

                tmze.dt1.Columns.Add("HenkanNouki");
                tmze.dt1.Columns.Add("ErrorKBN");
                tmze.dt1.Columns.Add("ErrorText");

                DataColumn dc1 = new DataColumn();
                dc1.DefaultValue = 0;
                dc1.ColumnName = "SEQ";
                tmze.dt1.Columns.Add(dc1);

                tmie = new T_MakerInport_Entity();
                tmie.DataSourseMakerCD = tmze.MakerCD;
                tmie.InportFile = tmze.ImportFileName;
                tmie.InsertOperator = tmze.InsertOperator;
                tmie.PatternCD = patternCD;
                tmie.xml = DataTableToXml(tmze.dt1);

                return psks0103idl.Import(tmie, insertDateTime,noukiFlag);
                
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool PSKS0103I_DELETE(T_MakerInport_Entity tmie)
        {
            return psks0103idl.PSKS0103I_DELETE(tmie);
        }

        public DataTable M_SiiresakiZaiko_Select(M_SiiresakiZaiko_Entity MSiireiSaki_Data)
        {
            return mszdl.M_SiiresakiZaiko_Select(MSiireiSaki_Data);
        }

        public DataTable T_MakerInport_Select(string InportSEQ, string ErrChkFlg)
        {
            return tmidl.T_MakerInport_Select(InportSEQ, ErrChkFlg);
        }

        public string JANCD_PositionCheck(string MakerCD)
        {
            return psks0103idl.JANCD_PositionCheck(MakerCD);
        }

        public bool MakerBrandSelect(string fields, string tableName, string condition)
        {
            DataTable dt = new DataTable();
            dt = bdl.DynamicSelectData(fields, tableName, condition);
            if (dt.Rows.Count > 0)
                return true;
            else return false;
        }

        public DataTable M_MakerZaiko_D_Select(MakerZaiko_D_Entity maker_D_entity)
        {
            return mmzddl.M_MakerZaiko_D_Select(maker_D_entity);
        }

        public bool M_SiiresakiZaiko_Update(M_SiiresakiZaiko_Entity msze)
        {
            return mszdl.M_SiiresakiZaiko_Update(msze);
        }
    }
}
