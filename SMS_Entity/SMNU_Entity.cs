using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace SMS_Entity
{
    public class SMNU_Entity : Base_Entity
    {
        public string MNUCD { get; set; }
        public string MNUNM { get; set; }
        public string GYMNO { get; set; }
        public string GYMCD { get; set; }
        public string PRONO { get; set; }
        public string PROID { get; set; }
        //public string MNUM1 { get; set; }
        //public string MNUM2 { get; set; }
        //public string MNUM3 { get; set; }
        //public string MNUV1 { get; set; }
        //public string MNUV2 { get; set; }
        //public string MNUV3 { get; set; }
        //public string ADDDT { get; set; }
        //public string ADDOP { get; set; }
        //public string UPDDT { get; set; }
        //public string UPDOP { get; set; }
        public DataTable dtSMNU { get; set; }
        public string xml { get; set; }
    }
}
