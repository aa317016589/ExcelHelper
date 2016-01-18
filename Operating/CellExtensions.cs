using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;


namespace ExcelHelper.Operating
{
    public static class CellExtensions
    {
        public static int MergeTo(this ICell self, ICell target)
        {
            int firstRow = self.RowIndex;
            int lastRow = target.RowIndex;
            int firstCol = self.ColumnIndex;
            int lastCol = target.ColumnIndex;

            if (self.RowIndex > target.RowIndex)
            {
                firstRow = target.RowIndex;
                lastRow = self.RowIndex;
            }

            if (self.ColumnIndex > target.ColumnIndex)
            {
                firstCol = target.ColumnIndex;
                lastCol = self.ColumnIndex;
            }

            CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            return self.Sheet.AddMergedRegion(cellRangeAddress);
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="width">一个单位为一个字符宽度</param>
        public static void SetColumnWidth(this ICell cell, int width)
        {
            cell.Sheet.SetColumnWidth(cell.ColumnIndex, (int)((width + 0.72) * 256));
        }

        /// <summary>
        /// 设置行高
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="height">一个单位为一个点长度</param>
        public static void SetRowHeight(this ICell cell, short height)
        {
            cell.Row.Height = (short)(height * 20);
        }
    }
}
