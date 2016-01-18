using ExcelHelper.Operating;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestProject.Model
{
    public class SheetDetailDataWrapper : ISheetDataWrapper
    {
        public virtual String DataName { get; set; }

        public virtual IEnumerable<IExcelModelBase> Datas { get; set; }

        public Int32 EmptyIntervalRow { get; set; }

        public Int32 CellWidth = 15;

        public short CellHeight = 18;

        public Boolean HaveTitle { get; set; }

        public SheetDetailDataWrapper(String dataName = "", IEnumerable<IExcelModelBase> datas = null)
        {
            this.DataName = dataName;

            this.Datas = datas == null ? new List<IExcelModelBase>() : datas;

            this.EmptyIntervalRow = 2;

            this.HaveTitle = true;
        }



        private IEnumerable<IExtendedBase> _titles;

        public virtual IEnumerable<IExtendedBase> Titles
        {
            get
            {
                if (_titles == null)
                {
                    if (Datas == null || Datas.Count() == 0)
                    {
                        return new List<IExtendedBase>();
                    }

                    _titles = Datas.SelectMany(s => s.ExtendedCollection).Where(s => s.TypeId > 0 && !String.IsNullOrEmpty(s.TypeName)).Distinct(s => s.TypeId).ToList();
                }

                return _titles;

            }
            set { _titles = value; }
        }
    }

    public class SheetDetail : ISheetDetail
    {
        public virtual String SheetName { get; set; }

        public virtual Boolean IsContinue { get; set; }

        public virtual SheetDataCollection SheetDetailDataWrappers { get; set; }

        public SheetDetail(String sheetName, SheetDataCollection datas = null)
        {
            this.SheetName = sheetName;

            this.SheetDetailDataWrappers = datas;

            this.IsContinue = false;

            this.SheetDetailDataWrappers = new SheetDataCollection();
        }

        public SheetDetail(String sheetName, ISheetDataWrapper datas)
            : this(sheetName, new SheetDataCollection())
        {
            this.SheetDetailDataWrappers.Add(datas);
        }
    }
}
