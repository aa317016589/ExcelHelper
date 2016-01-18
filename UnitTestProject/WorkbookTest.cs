using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using UnitTestProject.Model;
using ExcelHelper.Operating;
using System.Linq;
using System.IO;

namespace UnitTestProject
{
    [TestClass]
    public class WorkbookTest
    {
        [TestMethod]
        public void WorkbookInsert()
        {
            List<SumAnalysisX> SumAnalysisXs = new List<SumAnalysisX>();

            for (int i = 0; i < 8; i++)
            {
                SumAnalysisX SumAnalysisX1 = new SumAnalysisX() { ClassName = i.ToString() + "班级", MaxScore = 100 + i, Total = 1000 + i, TotalAverage = 10 + i };

                List<SubjectScoreDetailX> ssd = new List<SubjectScoreDetailX>();

                for (int j = 0; j < 4; j++)
                {
                    ssd.Add(new SubjectScoreDetailX() { Score = 10000 + j, TypeName = "kemu " + j.ToString(), TypeId = j });
                }

                SumAnalysisX1.SubjectScoreDetails = ssd;

                SumAnalysisXs.Add(SumAnalysisX1);

            }

            WorkbookWrapper w = new WorkbookWrapper();

            SheetDetail sd = new SheetDetail("sxf");

            SheetDetailDataWrapper sheetDetailDataWrapper = new SheetDetailDataWrapper("No.1", SumAnalysisXs.Cast<IExcelModelBase>());

            sd.SheetDetailDataWrappers.Add(sheetDetailDataWrapper);

            w.AddSheetDetail(sd);

            w.Save(@"C:\sxf.xls");
        }
        [TestMethod]
        public void WorkbookInsert2()
        {
            String path = AppDomain.CurrentDomain.BaseDirectory + "/D.xls";

            IEnumerable<PointsCoupon> pcs;

            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                 pcs = WorkbookBuilder.ExcelToDataTable<PointsCoupon>(fs, false);             
            }

            foreach (var item in pcs)
            {
                Console.WriteLine(item.ParValue);
            }
        }
    }
}