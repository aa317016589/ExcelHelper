using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Reflection;
using System.Web;

namespace ExcelHelper.Operating
{
    public class EnumService
    {
        public static string GetDescription(Enum obj)
        {
            if (obj == null)
            {
                return String.Empty;
            }

            string objName = obj.ToString();
            Type t = obj.GetType();
            FieldInfo fi = t.GetField(objName);
            DescriptionAttribute[] arrDesc = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (arrDesc.Length == 0)
            {
                return obj.ToString();
            }

            return arrDesc[0].Description;
        }
    }
}