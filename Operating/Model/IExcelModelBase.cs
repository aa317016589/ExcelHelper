 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper.Operating
{
    public interface IExcelModelBase
    {
        IEnumerable<IExtendedBase> ExtendedCollection { get; }
    }
}