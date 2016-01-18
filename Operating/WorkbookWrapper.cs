using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Operating
{
    /// <summary>
    /// excel容器，一个实例代表一个excel文件
    /// </summary>
    public class WorkbookWrapper
    {
        public WorkbookWrapper()
        {
            workbook = new HSSFWorkbookBuilder();

            SheetDetails = new List<ISheetDetail>();
        }


        public IList<ISheetDetail> SheetDetails;

        private HSSFWorkbookBuilder workbook;

        private Boolean IsFinish = false;

        public void AddSheetDetail(ISheetDetail sheetDetail)
        {
            SheetDetails.Add(sheetDetail);
        }

        private void Insert()
        {
            if (IsFinish)
            {
                return;
            }

            foreach (var item in SheetDetails)
            {
                workbook.Insert(item);
            }

            IsFinish = true;
        }


        public void Save(String path)
        {
            Insert();
 
            workbook.currentWorkbook.Save(path);
        }

        public void Download(String fileName)
        {
            Insert();

            workbook.currentWorkbook.Download(fileName);
        }
    }
}