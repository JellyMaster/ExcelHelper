using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    [Obsolete("Use the newer method of ExcelOutputBinding Instead")]
    public class SuppressFromExcel : Attribute
    {
        public bool Suppress { get; set; }


        public SuppressFromExcel(bool suppress = true)
        {
            Suppress = suppress;
        }

    }
}
