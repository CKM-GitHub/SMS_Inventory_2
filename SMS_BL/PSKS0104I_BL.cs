using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;

namespace SMS_BL
{
    public class PSKS0104I_BL : Base_BL
    {
        M_SiiresakiZaiko_DL mszdl;
        ItemSearch_DL isdl;
        T_MakerZaiko_DL tmzdl;
        public PSKS0104I_BL()
        {
            mszdl = new M_SiiresakiZaiko_DL();
            isdl = new ItemSearch_DL();
            tmzdl = new T_MakerZaiko_DL();
        }

        public DataTable M_SiiresakiZaiko_Select(M_SiiresakiZaiko_Entity msze)
        {
            return mszdl.M_SiiresakiZaiko_Select(msze);
        }

        public DataTable PSKS0104I_Select(ItemSearch_Entity ise)
        {
            return isdl.PSKS0104I_Select(ise);
        }

        public DataTable JANCD_ExistsCheck(string JANCD)
        {
           // return DynamicSelectData("VC_JAN1,VM_SHOHIN", "M_KIHON_SHOHIN", "WHERE VC_JAN1 = '" + JANCD + "' and VC_SHIIRESAKI='" + MakerCD+"'");
            return DynamicSelectData("VC_JAN1,VM_SHOHIN", "M_KIHON_SHOHIN", "WHERE VC_JAN1 = '" + JANCD + "'");
        }

        public DataTable T_MakerZaiko_ExistCheck(T_MakerZaiko_Entity T_MakerZaikoData)
        {
            return DynamicSelectData("JanCD", "T_MakerZaiko", "where MakerCD = '" + T_MakerZaikoData.MakerCD + "' and JanCD = '" + T_MakerZaikoData.JanCD + "'");
        }

        public bool psks0104I_InsertUpdate(T_MakerZaiko_Entity tmze, L_Log_Entity lle)
        {

            tmzdl.StartTransaction();
            tmze.InportDateTime = lle.OperateDate = System.DateTime.Now.ToString();
            if (tmze.dt1 != null)
            {
                tmze.dt1.Columns.Remove("vm_hanbai_shohin");
                tmze.xml = DataTableToXml(tmze.dt1);

                if (!tmzdl.T_MakerZaiko_Update(tmze))
                    return false;
            }
            M_SiiresakiZaiko_Entity msze = new M_SiiresakiZaiko_Entity();
            msze.SiiresakiCD = tmze.MakerCD;
            msze.KakuTeiFlg = tmze.KakuTeiFlg;
            msze.InsertDateTime = tmze.InsertDateTime;

            M_SiiresakiZaiko_DL mszdl = new M_SiiresakiZaiko_DL();
            L_Log_DL lldl = new L_Log_DL();

            mszdl.Transaction = lldl.Transaction = tmzdl.Transaction;

            if (mszdl.M_SiiresakiZaiko_Update(msze))
            {
                if (lldl.L_Log_Insert(lle))
                {
                    tmzdl.CommitTransaction();
                    return true;
                }
            }

            return false;
        }
    }
}
