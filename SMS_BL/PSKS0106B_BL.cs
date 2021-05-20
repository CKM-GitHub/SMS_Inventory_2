using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using SMS_Entity;
using System.Data;
using System.Collections;

namespace SMS_BL
{
    public class PSKS0106B_BL : Base_BL
    {
        M_MultiPorpose_DL mmpdl;
        T_Ukebarai_DL tudl;
        Base_DL BaseDL;

        public PSKS0106B_BL()
        {
            mmpdl = new M_MultiPorpose_DL();
            tudl = new T_Ukebarai_DL();
            BaseDL = new Base_DL();
        }

        public DataTable M_MultiPorpose_SelectID(M_MultiPorpose_Entity MmultiporposeData)
        {
            return BaseDL.DynamicSelectData("Char1 ,Char2,Num1,Num2", "M_MultiPorpose", "where ID= " + MmultiporposeData.ID + " AND [Key]='" + MmultiporposeData.Key + "'");
        }

        public DataTable M_MultiPorpose_SelectChar()
        {
            return BaseDL.DynamicSelectData("Char1", "M_MultiPorpose", "where ID=102 AND [Key]= '1' ");
        }

        public bool Ukebarai_Insert( DataTable dtCsvAll)
        {
            string CSV_xml = DataTableToXml(dtCsvAll);
            return tudl.PSKS0106B_Insert(CSV_xml);
        }

    }
}
