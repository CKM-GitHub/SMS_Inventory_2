using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_Entity 
{
    public class T_NyuukinMeisai_Entity : Base_Entity
    {
       public string OperatorCD { get; set; }
       public string DataKBN { get; set; }
       public string KouzaNyuukinDate { get; set; }
       public string NyuukinMotoCD { get; set; }
       public string DetailNo { get; set; }
       public string NyuukinMotoName { get; set; }
       public string KeijouDate { get; set; }
       public string NyuukinGaku { get; set; }
       public string JuShuuNo { get; set; }
       public string WebJuchuuNo { get; set; }
       public string NyuukinKBN { get; set; }
       public string KokyakuNo { get; set; }
       public string KokyakuName { get; set; }
       public string GenkinGaku { get; set; }
       public string KogitteGaku { get; set; }
       public string TegataGaku { get; set; }
       public string SousaiGaku { get; set; }
       public string TyouseiGaku { get; set; }
       public string NoUseFlag { get; set; }
       public string InsertDatetime { get; set; }
       public string RenkeiDatetime { get; set; }
        
    }
}
