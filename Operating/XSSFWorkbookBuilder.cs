using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelHelper.Operating
{
    /// <summary>
    /// Excel2007
    /// </summary>
    public class XSSFWorkbookBuilder : WorkbookBuilder
    {
        protected override IWorkbook CreateWorkbook()
        {
            this.OnContentCellSetAfter += sxf;

            return new XSSFWorkbook();
        }

        public void sxf(BuildContext context)
        {
            context.Cell.CellStyle = this.CenterStyle;
        }
    }
}
