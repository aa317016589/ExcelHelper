using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


namespace ExcelHelper.Operating
{
    public static class ExcelModelsPropertyManage
    {
        private static Dictionary<String, IList<PropertyInfoDetail>> PropertyInfos;

        private static Dictionary<String, Object> ExtendedDefaultValue;

        static ExcelModelsPropertyManage()
        {
            PropertyInfos = new Dictionary<String, IList<PropertyInfoDetail>>();

            ExtendedDefaultValue = new Dictionary<String, Object>();
        }

        #region 获取属性列表
        public static IList<PropertyInfoDetail> CreatePropertyInfos(Type t)
        {
            String name = t.Name;

            if (PropertyInfos.ContainsKey(name))
            {
                return PropertyInfos[name];
            }

            PropertyInfo[] ps = t.GetProperties();

            List<PropertyInfoDetail> checkPropertyInfos = new List<PropertyInfoDetail>();

            foreach (var item in ps)
            {
                Object[] d = item.GetCustomAttributes(false);

                if (d.Where(s => s.GetType() == typeof(IgnoreAttribute)).Any())
                {
                    continue;
                }

                DefaultValueAttribute dv = d.Where(s => s.GetType() == typeof(DefaultValueAttribute)).SingleOrDefault() as DefaultValueAttribute;

                checkPropertyInfos.Add(new PropertyInfoDetail() { PropertyInfoV = item, DefaultVale = dv == null ? null : dv.Value });
            }

            PropertyInfos.Add(name, checkPropertyInfos);

            return checkPropertyInfos;
        }
        #endregion

        public static Object GetExtendedDefaultValue(Type t)
        {
            String name = t.Name;

            if (ExtendedDefaultValue.ContainsKey(name))
            {
                return ExtendedDefaultValue[name];
            }



            PropertyInfo p = t.GetProperties().Where(s => s.Name == "TypeValue").SingleOrDefault();
             
            if (p == null)
            {
                return null;
            }

            DefaultValueAttribute dv = p.GetCustomAttributes(typeof(DefaultValueAttribute), false).SingleOrDefault() as DefaultValueAttribute;

            Object v = dv == null ? null : dv.Value;

            ExtendedDefaultValue.Add(name, v);

            return v;


        }
    }


    public class PropertyInfoDetail
    {
        public PropertyInfo PropertyInfoV { get; set; }

        public Object DefaultVale { set; get; }
    }
}