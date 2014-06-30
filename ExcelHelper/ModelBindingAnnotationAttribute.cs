using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public class ModelBindingAnnotationAttribute : Attribute
    {

        public string DBColumnName { get; set; }

        public ModelBindingAnnotationAttribute(string databaseColumnName = "")
        {
            DBColumnName = databaseColumnName;
        }

    }
}
