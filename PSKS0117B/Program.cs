using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_BL;
using SMS_Entity;
using System.Threading;
using System.IO;
using System.Collections;
using System.Data;

namespace PSKS0117B
{

    class Program
    {
        public static void Main(string[] args)
        {
            Console.Title = "117B";
            PSKS0117B_BL psks0117b_bl = new PSKS0117B_BL();

            bool result = false;
            result= psks0117b_bl.Log_Delete();


        }
    }
}
