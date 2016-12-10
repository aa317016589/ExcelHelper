using ExcelHelper.Operating;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHelper.Operating.Model;

namespace UnitTestProject.Model
{
    public class SheetDetailDataWrapper : ISheetDataWrapper
    {
        public String DataName { get; set; }

        public IEnumerable<IExcelModelBase> Datas { get; set; }

        public Int32 EmptyIntervalRow { get; set; }

        public Int32 CellWidth = 15;

        public short CellHeight = 18;

        public Boolean HaveTitle { get; set; }

        public SheetDetailDataWrapper(String dataName = "", IEnumerable<IExcelModelBase> datas = null)
        {
            this.DataName = dataName;

            this.Datas = datas ?? new List<IExcelModelBase>();

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
                    if (Datas == null || !Datas.Any())
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
        public String SheetName { get; set; }

        public Boolean IsContinue { get; set; }

        public SheetDataCollection SheetDetailDataWrappers { get; set; }

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
