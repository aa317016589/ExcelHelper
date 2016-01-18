using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ExcelHelper.Operating
{
    /// <summary>
    /// Excel2003
    /// </summary>
    public class HSSFWorkbookBuilder : WorkbookBuilder
    {

        protected override IWorkbook CreateWorkbook()
        {
            this.OnContentCellSetAfter += SetStyle;

            return new HSSFWorkbook();
        }

        public void SetStyle(BuildContext context)
        {
            context.Cell.SetRowHeight(18);

            //context.Cell.SetColumnWidth(15);

            context.Cell.CellStyle = this.CenterStyle;
        }
    }
}