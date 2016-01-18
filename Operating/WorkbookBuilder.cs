using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ComponentModel;
using System.Collections;


namespace ExcelHelper.Operating
{
    public abstract class WorkbookBuilder
    {
        protected WorkbookBuilder()
        {
            currentWorkbook = CreateWorkbook();

            buildContext = new BuildContext() { WorkbookBuilder = this, Workbook = currentWorkbook };
        }

        public delegate void BuildEventHandler(BuildContext context);

        protected abstract IWorkbook CreateWorkbook();

        public IWorkbook currentWorkbook;

        private ICellStyle _centerStyle;

        public ICellStyle CenterStyle
        {
            get
            {
                if (_centerStyle == null)
                {
                    _centerStyle = currentWorkbook.CreateCellStyle();

                    _centerStyle.Alignment = HorizontalAlignment.Center;

                    _centerStyle.VerticalAlignment = VerticalAlignment.Center;
                }

                return _centerStyle;
            }
        }

        private Int32 StartRow = 0;//起始行


        private BuildContext buildContext;
 
        public event BuildEventHandler OnHeadCellSetAfter;
 
        public event BuildEventHandler OnContentCellSetAfter;


        #region DataTableToExcel

        public void Insert(ISheetDetail sheetDetail)
        {
            ISheet sheet;

            if (sheetDetail.IsContinue)
            {
                sheet = currentWorkbook.GetSheetAt(currentWorkbook.NumberOfSheets - 1);

                StartRow = sheet.LastRowNum + 1;
            }
            else
            {
                sheet = currentWorkbook.CreateSheet(sheetDetail.SheetName);
            }

            buildContext.Sheet = sheet;

            sheet = DataToSheet(sheetDetail.SheetDetailDataWrappers, sheet);

        }
        /// <summary>
        /// 这里添加数据，循环添加，主要应对由多个组成的
        /// </summary>
        /// <param name="sheetDetailDataWrappers"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private ISheet DataToSheet(SheetDataCollection sheetDetailDataWrappers, ISheet sheet)
        {
            foreach (var sheetDetailDataWrapper in sheetDetailDataWrappers)
            {
                if (sheetDetailDataWrapper.Datas == null || sheetDetailDataWrapper.Datas.Count() == 0)
                {
                    continue;
                }

                Type type = sheetDetailDataWrapper.Datas.GetType().GetGenericArguments()[0];

                if (sheetDetailDataWrapper.HaveTitle)
                {
                    sheet = SetTitle(sheet, sheetDetailDataWrapper, type);
                }

                sheet = AddValue(sheet, sheetDetailDataWrapper, type);

                StartRow = StartRow + sheetDetailDataWrapper.EmptyIntervalRow;
            }

            return sheet;
        }

        #endregion

        #region 设置值

        private void SetCellValue(ICell cell, object obj)
        {
            if (obj == null)
            {
                cell.SetCellValue(" "); return;
            }
  
            if (obj is String)
            {
                cell.SetCellValue(obj.ToString()); return;
            }

            if (obj is Int32 || obj is Double)
            {
                cell.SetCellValue(Math.Round(Double.Parse(obj.ToString()), 2)); return;
            }

            if (obj.GetType().IsEnum)
            {
                cell.SetCellValue(EnumService.GetDescription((Enum)obj)); return;
            }

            if (obj is DateTime)
            {
                cell.SetCellValue(((DateTime)obj).ToString("yyyy-MM-dd HH:mm:ss")); return;
            }

            if (obj is Boolean)
            {
                cell.SetCellValue((Boolean)obj ? "√" : "×"); return;
            }     
        }

        #endregion

        #region SetTitle
        private ISheet SetTitle(ISheet sheet, ISheetDataWrapper sheetDetailDataWrapper, Type type)
        {
            IRow titleRow = null;

            ICell titleCell = null;

            if (!String.IsNullOrEmpty(sheetDetailDataWrapper.DataName))
            {
                titleRow = sheet.CreateRow(StartRow);

                buildContext.Row = titleRow;
 
                StartRow++;

                titleCell = SetCell(titleRow, 0, sheetDetailDataWrapper.DataName);

                if (OnHeadCellSetAfter != null)
                {
                    OnHeadCellSetAfter(buildContext);
                }
            }

            IRow row = sheet.CreateRow(StartRow);

            buildContext.Row = row;

            IList<PropertyInfo> checkPropertyInfos = ExcelModelsPropertyManage.CreatePropertyInfos(type);

            int i = 0;

            foreach (PropertyInfo property in checkPropertyInfos)
            {
                DisplayNameAttribute dn = property.GetCustomAttributes(typeof(DisplayNameAttribute), false).SingleOrDefault() as DisplayNameAttribute;

                if (dn != null)
                {
                    SetCell(row, i++, dn.DisplayName);
                    continue;
                }

                Type t = property.PropertyType;

                if (t.IsGenericType)
                {
                    if (sheetDetailDataWrapper.Titles == null || sheetDetailDataWrapper.Titles.Count() == 0)
                    {
                        continue;
                    }

                    foreach (var item in sheetDetailDataWrapper.Titles)
                    {
                        SetCell(row, i++, item.TypeName);
                    }
                }
            }
        
            if (titleCell != null && i > 0)
            {
                titleCell.MergeTo(titleRow.CreateCell(i - 1));

                titleCell.CellStyle = this.CenterStyle;
            }

            StartRow++;

            return sheet;
        }
        #endregion

        #region AddValue
        private ISheet AddValue(ISheet sheet, ISheetDataWrapper sheetDetailDataWrapper, Type type)
        {
            IList<PropertyInfo> checkPropertyInfos = ExcelModelsPropertyManage.CreatePropertyInfos(type);

            Int32 cellCount = 0;

            foreach (var item in sheetDetailDataWrapper.Datas)
            {
                if (item == null)
                {
                    StartRow++;
                    continue;
                }

                IRow newRow = sheet.CreateRow(StartRow);

                buildContext.Row = newRow;

                foreach (PropertyInfo property in checkPropertyInfos)
                {
                    Object obj = property.GetValue(item, null);

                    Type t = property.PropertyType;

                    ICell cell;

                    if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                    {
                        var ssd = ((IEnumerable)obj).Cast<IExtendedBase>();

                        if (ssd == null)
                        {
                            continue;
                        }

                        foreach (var v in sheetDetailDataWrapper.Titles)
                        {
                            IExtendedBase sv = ssd.Where(s => s.TypeId == v.TypeId).SingleOrDefault();

                            SetCell(newRow, cellCount++, sv.TypeValue);
                        }

                        continue;
                    }
 
                    SetCell(newRow, cellCount++, obj);
                }

                StartRow++;
                cellCount = 0;
            }

            return sheet;
        }

        #endregion

        #region 设置单元格
        /// <summary>
        /// 设置单元格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="index"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private ICell SetCell(IRow row, int index, object value)
        {
            ICell cell = row.CreateCell(index);

            SetCellValue(cell, value);

            buildContext.Cell = cell;

            if (OnContentCellSetAfter != null)
            {
                OnContentCellSetAfter(buildContext);
            }

            return cell;
        } 
        #endregion

        #region ExcelToDataTable

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">具体对象</typeparam>
        /// <param name="fs"></param>
        /// <param name="fileName"></param>
        /// <param name="isFirstRowColumn"></param>
        /// <returns></returns>
        public static IEnumerable<T> ExcelToDataTable<T>(Stream fs, bool isFirstRowColumn = false) where T : new()
        {
            List<T> ts = new List<T>();

            Type type = typeof(T);

            IList<PropertyInfo> checkPropertyInfos = ExcelModelsPropertyManage.CreatePropertyInfos(type);

            try
            {
                IWorkbook workbook = WorkbookFactory.Create(fs);

                fs.Dispose();

                ISheet sheet = workbook.GetSheetAt(0);

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);

                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    Int32 startRow = isFirstRowColumn ? 1 : 0;

                    int rowCount = sheet.LastRowNum; //行数

                    int length = checkPropertyInfos.Count;

                    length = length > cellCount + 1 ? cellCount + 1 : length;

                    Boolean haveValue = false;

                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);

                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        T t = new T();

                        for (int f = 0; f < length; f++)
                        {
                            ICell cell = row.GetCell(f);

                            if (cell == null || String.IsNullOrEmpty(cell.ToString()))
                            {
                                continue;
                            }

                            object b = cell.ToString();

                            if (cell.CellType == CellType.Numeric)
                            {
                                //NPOI中数字和日期都是NUMERIC类型的，这里对其进行判断是否是日期类型
                                if (HSSFDateUtil.IsCellDateFormatted(cell))//日期类型
                                {
                                    b = cell.DateCellValue;
                                }
                                else
                                {
                                    b = cell.NumericCellValue;
                                }
                            }

                            PropertyInfo pinfo = checkPropertyInfos[f];

                            if (pinfo.PropertyType.Name != b.GetType().Name) //类型不一样的时候，强转
                            {
                                b = System.ComponentModel.TypeDescriptor.GetConverter(pinfo.PropertyType).ConvertFrom(b.ToString());
                            }

                            type.GetProperty(pinfo.Name).SetValue(t, b, null);

                            if (!haveValue)
                            {
                                haveValue = true;
                            }
                        }
                        if (haveValue)
                        {
                            ts.Add(t); haveValue = false;
                        }
                    }
                }

                return ts;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        #endregion
    }

    public class BuildContext
    {
        public WorkbookBuilder WorkbookBuilder { get; set; }
        
        public IWorkbook Workbook { get; set; }

        public ISheet Sheet { get; set; }

        public IRow Row { get; set; }

        public ICell Cell { get; set; }

    }
}