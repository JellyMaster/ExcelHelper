using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public class ExcelOutputBinding : Attribute
    {

 


        public bool Suppress { get; set; }

        public double Order { get; set; }

        public string Name { get; set; }

        public bool Hide { get; set; }


 


        public ExcelOutputBinding()
        {
            Suppress = false;
            Order = -1;
            Name = string.Empty;
            Hide = false;
            
        }




    }
}
