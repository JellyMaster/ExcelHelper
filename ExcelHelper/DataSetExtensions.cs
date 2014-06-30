using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public static class DataTableExtensions
    {

        /// <summary>
        /// This will flatten out a list of objects into a data table. 
        /// Current version will only handle standard data types within an object. 
        /// Will not flatten out a collection of objects if assigned as a property. eg List<string>
        /// </summary>
        /// <typeparam name="T">The object type we are trying to flatten</typeparam>
        /// <param name="items">The collection we are sending in. </param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            List<PropertyInfo> FilterPropertyInfo = new List<PropertyInfo>();
            List<KeyValuePair<int, PropertyInfo>> propertyPairs = new List<KeyValuePair<int, PropertyInfo>>();

            foreach (PropertyInfo prop in properties)
            {
                //Setting column names as Property names

                if (!ExcelOutputSuppress(prop))
                {

                    propertyPairs.Add(new KeyValuePair<int, PropertyInfo>(ExcelOutputOrder(prop), prop));

                }
            }


            //we know the list of columns hopefully 
            //now lets order them. 

            //simple ordering here. 
            FilterPropertyInfo = propertyPairs.OrderBy(s => s.Key).Select(s => s.Value).ToList();

            foreach (PropertyInfo prop in FilterPropertyInfo)
            {
                bool isNullable = false;
                dataTable.Columns.Add(new DataColumn(ExcelColumnName(prop), SafeCastConvertors.GetType(prop, out isNullable)));
            }





            foreach (T item in items)
            {
                var values = new object[dataTable.Columns.Count];

                int counter = 0;

                //as we know these are order we can just cycle through and add in the correct order.
                //because we have already figured out if we have supressed the column we don't need to check again. 


                foreach (PropertyInfo prop in FilterPropertyInfo)
                {
                    values[counter] = prop.GetValue(item, null);
                    counter++;
                }



                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

        public static Type ExcelDataType(PropertyInfo property)
        {
            bool isNullable = false;
            //we will return what ever type is set for the system. 
            Type type = SafeCastConvertors.GetType(property, out isNullable);

#if DEBUG

            Debug.WriteLine(string.Format("The property {0} is of type {1}", property.Name, property.PropertyType));

#endif




            return type;
        }



        /// <summary>
        /// Convert a datatable to a list collection of type T 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        /// <returns></returns>
        public static List<T> ToList<T>(this DataTable table)
        {
            List<T> model = new List<T>();

            //Get an array of the properties for the type of T 
            PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (DataRow row in table.Rows)
            {
                //Create a new instance of the type T 
                T item = Activator.CreateInstance<T>();



                //now iterate through the columns in the dataset and convert to the appropriate type for me. 
                foreach (DataColumn column in table.Columns)
                {
                    //as we do not know if the item is using the property name, display name or model binding name I will have to check for each
                    //attribute type as a check
                    foreach (PropertyInfo propInfo in properties)
                    {
                        if (FoundProperty(propInfo, column))
                        {



                            try
                            {
                                object value = row[column];

                                bool isNullable = false;


                                Type typeObject = SafeCastConvertors.GetType(propInfo, out isNullable);


                                switch (typeObject.FullName)
                                {
                                    case "System.String":
                                        {
                                            propInfo.SetValue(item, SafeCastConvertors.ParseType<string>(value));
                                            break;
                                        }
                                    case "System.Int32":
                                        {

                                            if (isNullable)
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<int?>(value));
                                            }
                                            else
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<int>(value));
                                            }

                                            break;
                                        }
                                    case "System.Double":
                                        {

                                            if (isNullable)
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<double?>(value));
                                            }
                                            else
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<double>(value));
                                            }

                                            break;
                                        }
                                    case "System.DateTime":
                                        {
                                            if (isNullable)
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<DateTime?>(value));
                                            }
                                            else
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<DateTime>(value));
                                            }
                                            break;
                                        }
                                    case "System.Boolean":
                                        {

                                            if (isNullable)
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<bool?>(value));
                                            }
                                            else
                                            {
                                                propInfo.SetValue(item, SafeCastConvertors.ParseType<bool>(value));
                                            }

                                            break;
                                        }
                                    default:
                                        {

                                            break;
                                        }
                                }

                            }
                            catch (Exception error)
                            {
                                Debug.WriteLine(error.Message);
                            }
                            break;
                        }
                    }

                }

                model.Add(item);

            }

            return model;
        }






        public static bool SuppressFromExcel(PropertyInfo property)
        {
            bool model = false;


            var suppress = property.GetCustomAttribute<SuppressFromExcel>();

            if (suppress != null && suppress.Suppress)
            {
                model = true;
            }


            return model;
        }


        public static bool ExcelOutputSuppress(PropertyInfo property)
        {
            bool model = false;


            var suppress = property.GetCustomAttribute<ExcelOutputBinding>();

            if (suppress != null && suppress.Suppress)
            {
                model = true;
            }
            else
            {
                //as a safety net check if SUppressFromExcel is there. 
                model = SuppressFromExcel(property);
            }


            return model;
        }



        public static int ExcelOutputOrder(PropertyInfo property)
        {
            int model = -1;
            var order = property.GetCustomAttribute<ExcelOutputBinding>();


            if (order != null && order.Order > -1)
            {
                model = order.Order;
            }
            else
            {
                //set as max value to push the values to the end of the collection 
                //setting to a minus figure will shift these to the start of the list
                //not what we want. 
                model = int.MaxValue;
            }

            return model;

        }



        public static string ExcelColumnName(PropertyInfo property)
        {
            string model = string.Empty;



            var displayName = property.GetCustomAttribute<ExcelOutputBinding>();


            if (displayName != null && !string.IsNullOrEmpty(displayName.Name))
            {
                model = displayName.Name;
            }
            else
            {
                model = ColumnName(property);
            }


            return model;
        }



        public static string ColumnName(PropertyInfo property)
        {
            string model = string.Empty;



            var displayName = property.GetCustomAttribute<DisplayAttribute>();


            if (displayName != null && !string.IsNullOrEmpty(displayName.Name))
            {
                model = displayName.Name;
            }
            else
            {
                model = property.Name;
            }


            return model;
        }





        public static bool FoundProperty(PropertyInfo property, DataColumn column)
        {
            bool valid = false;

            string columnName = column.ColumnName;

            if (property.Name == columnName)
            {
                valid = true;
            }
            else
            {
                //check to see if the display attribute is set for the column
                var displayAttribute = property.GetCustomAttribute<DisplayAttribute>();

                if (displayAttribute != null && columnName == displayAttribute.Name)
                {
                    valid = true;
                }
                else
                {
                    //final check is to see if the property is using the custom model binding attribute that I have in the system. 
                    var modelBindAttribute = property.GetCustomAttribute<ModelBindingAnnotationAttribute>();


                    if (modelBindAttribute != null && modelBindAttribute.DBColumnName == columnName)
                    {
                        valid = true;
                    }


                }

            }
            return valid;

        }

    }
}
