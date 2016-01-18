using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Operating
{
    public static class ExcelModelsPropertyManage
    {
        private static Dictionary<String, IList<PropertyInfo>> PropertyInfos;

        static ExcelModelsPropertyManage()
        {
            PropertyInfos = new Dictionary<string, IList<PropertyInfo>>();
        }

        #region 获取属性列表
        public static IList<PropertyInfo> CreatePropertyInfos(Type t)
        {
            String name = t.Name;

            if (PropertyInfos.ContainsKey(name))
            {
                return PropertyInfos[name];
            }

            PropertyInfo[] ps = t.GetProperties();

            List<PropertyInfo> checkPropertyInfos = new List<PropertyInfo>();

            foreach (var item in ps)
            {
                object[] os = item.GetCustomAttributes(typeof(IgnoreAttribute), false);

                if (os.Length > 0)
                {
                    continue;
                }

                checkPropertyInfos.Add(item);
            }

            PropertyInfos.Add(name, checkPropertyInfos);

            return checkPropertyInfos;
        }
        #endregion
    }
}
