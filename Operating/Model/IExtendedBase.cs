using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelHelper.Operating
{
   public interface IExtendedBase
    {
        Object TypeValue { get; set; }

        String TypeName { get; set; }

        Int32 TypeId { get; set; }
    }
}