using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public static class SafeCastConvertors
    {

        public static T ParseType<T>(object value)
        {
            try
            {
                if (typeof(T) != typeof(string))
                {
                    //create a new instance of the type 
                    T model = Activator.CreateInstance<T>();


                    bool isNullable = false;
                    PropertyInfo propInfo = value.GetType().GetProperties()[0];

                    string type = GetType(propInfo, out isNullable).FullName;

                    Debug.WriteLine(string.Format("From ParseType: Property Type is: {0}, value passed in is {1}", type, value.ToString()));


                    switch (type)
                    {
                        case "System.Int32":
                            {

                                int val;

                                if (int.TryParse(value.ToString(), out val))
                                {
                                    model = (T)(object)val;
                                }

                                break;
                            }
                        case "System.DateTime":
                            {
                                DateTime val;

                                if (DateTime.TryParse(value.ToString(), out val))
                                {
                                    model = (T)(object)val;
                                }
                                break;
                            }
                        case "System.Double":
                            {
                                double val;

                                if (double.TryParse(value.ToString(), out val))
                                {
                                    model = (T)(object)val;
                                }
                                break;
                            }
                        case "System.Nullable":
                            {

                                break;
                            }
                        case "System.Boolean":
                            {
                                bool val;
                                if (bool.TryParse(value.ToString(), out val))
                                {
                                    model = (T)(object)val;
                                }
                                break;
                            }
                        default:
                            {
                                break;
                            }

                    }



                    return model;
                }
                else
                {
                    //return string value as is 
                    return (T)value;
                }
            }
            catch (Exception error)
            {
                Debug.WriteLine(error.Message);

                return Activator.CreateInstance<T>();
            }
        }




        public static Type GetType(PropertyInfo property, out bool isNullable)
        {
            isNullable = false;
            Type returnType = property.PropertyType;

            //test if we have a nullable type here. 
            if (returnType.IsGenericType && returnType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                //return the first type from the list. This really should only be one value. 
                //haven't seen an instance yet where it is more than one 
                returnType = returnType.GenericTypeArguments[0];
                isNullable = true;
            }

#if DEBUG
            Debug.WriteLine(string.Format("The property {0} is of return type {1} and Nullable Status is {2}", property.Name, returnType.Name, isNullable));
            Debug.WriteLine(string.Format("The property {0} is of type {1} and Nullable Status is {2}", property.Name, property.PropertyType, isNullable));
#endif
            return returnType;
        }



    }
}
