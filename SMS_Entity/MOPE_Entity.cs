using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace SMS_Entity
{
    public class MOPE_Entity : Base_Entity
    {
        public string TANCD { get; set; }
        public string TANNM { get; set; }
        public string TANRN { get; set; }
        public string TANKN { get; set; }
        public string ETANN { get; set; }
        public string ETANR { get; set; }
        public string BMNCD { get; set; }
        public string TMPCD { get; set; }
        public string PASWD { get; set; }
        public string MNUCD { get; set; }
        public string DELFG { get; set; }
        public string ADDDT { get; set; }
        public string ADDOP { get; set; }
        public string UPDDT { get; set; }
        public string UPDOP { get; set; }
        public string DELDT { get; set; }
        public string DELOP { get; set; }
        public string KNGCD { get; set; }
        public string GYMNM { get; set; }
        public string PRONM { get; set; }
        public string PRONO { get; set; }
        public string GYMNO { get; set; }
        public string PROID { get; set; }
        public DataTable dtMope { get; set; }
    }
}
