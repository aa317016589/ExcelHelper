using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Operating
{
    /// <summary>
    /// 表内部的数据容器
    /// </summary>
    public class SheetDataCollection : ICollection<ISheetDataWrapper>
    {
        IList<ISheetDataWrapper> SheetDetailDataWrappers;
 
        public SheetDataCollection()
        {
            this.SheetDetailDataWrappers = new List<ISheetDataWrapper>();
        }

        public void Add(ISheetDataWrapper item)
        {
            SheetDetailDataWrappers.Add(item);
        }

        public bool Contains(ISheetDataWrapper item)
        {
            return SheetDetailDataWrappers.Contains(item);
        }

        public void CopyTo(ISheetDataWrapper[] array, int arrayIndex)
        {
            SheetDetailDataWrappers.CopyTo(array, arrayIndex);
        }

        public bool Remove(ISheetDataWrapper item)
        {
            return SheetDetailDataWrappers.Remove(item);
        }

        IEnumerator<ISheetDataWrapper> IEnumerable<ISheetDataWrapper>.GetEnumerator()
        {
            return SheetDetailDataWrappers.GetEnumerator();
        }


        public void Clear()
        {
            SheetDetailDataWrappers.Clear();
        }

        public int Count
        {
            get { return SheetDetailDataWrappers.Count; }
        }

        public bool IsReadOnly
        {
            get { return SheetDetailDataWrappers.IsReadOnly; }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return SheetDetailDataWrappers.GetEnumerator();
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