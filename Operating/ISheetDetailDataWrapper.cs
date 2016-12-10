using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHelper.Operating.Model;

namespace ExcelHelper.Operating
{
    /// <summary>
    /// 表内部的数据容器
    /// </summary>
    public class SheetDataCollection : ICollection<ISheetDataWrapper>
    {
        private readonly IList<ISheetDataWrapper> _sheetDetailDataWrappers;

        public SheetDataCollection()
        {
            this._sheetDetailDataWrappers = new List<ISheetDataWrapper>();
        }

        public void Add(ISheetDataWrapper item)
        {
            _sheetDetailDataWrappers.Add(item);
        }

        public bool Contains(ISheetDataWrapper item)
        {
            return _sheetDetailDataWrappers.Contains(item);
        }

        public void CopyTo(ISheetDataWrapper[] array, int arrayIndex)
        {
            _sheetDetailDataWrappers.CopyTo(array, arrayIndex);
        }

        public bool Remove(ISheetDataWrapper item)
        {
            return _sheetDetailDataWrappers.Remove(item);
        }

        IEnumerator<ISheetDataWrapper> IEnumerable<ISheetDataWrapper>.GetEnumerator()
        {
            return _sheetDetailDataWrappers.GetEnumerator();
        }


        public void Clear()
        {
            _sheetDetailDataWrappers.Clear();
        }

        public int Count
        {
            get { return _sheetDetailDataWrappers.Count; }
        }

        public bool IsReadOnly
        {
            get { return _sheetDetailDataWrappers.IsReadOnly; }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _sheetDetailDataWrappers.GetEnumerator();
        }
    }


    public interface ISheetDataWrapper
    {
        String DataName { get; set; }

        Int32 EmptyIntervalRow { get; set; }

        Boolean HaveTitle { get; set; }

        IEnumerable<IExcelModelBase> Datas { get; }

        IEnumerable<IExtendedBase> Titles { get; set; }
    }
}