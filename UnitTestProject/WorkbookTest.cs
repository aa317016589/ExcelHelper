using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using UnitTestProject.Model;
using ExcelHelper.Operating;
using System.Linq;
using System.IO;
using ExcelHelper.Operating.Model;

namespace UnitTestProject
{
    [TestClass]
    public class WorkbookTest
    {
        [TestMethod]
        public void WorkbookInsert()
        {
            List<SumAnalysisX> sumAnalysisXs = new List<SumAnalysisX>();

            for (int i = 0; i < 8; i++)
            {
                SumAnalysisX sumAnalysisX1 = new SumAnalysisX() { ClassName = null, MaxScore = 100 + i, Total = 1000 + i, TotalAverage = 10 + i };

                List<SubjectScoreDetailX> ssd = new List<SubjectScoreDetailX>();

                for (int j = 0; j < 4; j++)
                {
                    ssd.Add(new SubjectScoreDetailX() { Score = null, TypeName = "kemu " + j.ToString(), TypeId = j });
                }

                sumAnalysisX1.SubjectScoreDetails = ssd;

                sumAnalysisXs.Add(sumAnalysisX1);

            }

            WorkbookWrapper w = new WorkbookWrapper();

            SheetDetail sd = new SheetDetail("sxf");

            SheetDetailDataWrapper sheetDetailDataWrapper = new SheetDetailDataWrapper("No.1", sumAnalysisXs.Cast<IExcelModelBase>());

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