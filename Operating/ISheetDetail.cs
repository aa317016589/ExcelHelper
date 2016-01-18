using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelHelper.Operating
{
    /// <summary>
    /// 
    /// </summary>
    public interface ISheetDetail
    {
        String SheetName { get; set; }

        Boolean IsContinue { get; set; }

        SheetDataCollection SheetDetailDataWrappers { get; set; }
    }
}