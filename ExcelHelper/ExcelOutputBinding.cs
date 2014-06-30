using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public class ExcelOutputBinding : Attribute
    {

        public enum CellFormat
        {
            String,
            Numeric,
            Date,
            Time,
            DateTime
        }


        public bool Suppress { get; set; }

        public int Order { get; set; }

        public string Name { get; set; }

        public bool Hide { get; set; }


        public CellFormat Format { get; set; }


        public ExcelOutputBinding()
        {
            Suppress = false;
            Order = -1;
            Name = string.Empty;
            Hide = false;
            Format = CellFormat.String;
        }




    }
}
