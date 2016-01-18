using ExcelHelper.Operating;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestProject.Model
{
    public class SumAnalysisX : IExcelModelBase
    {
        [DisplayName("班级")]
        public String ClassName { get; set; }

        /// <summary>
        /// 最高分
        /// </summary>
        [DisplayName("最高分")]
        public Double MaxScore { get; set; }


        public IEnumerable<SubjectScoreDetailX> SubjectScoreDetails { get; set; }


        [Ignore]
        public IEnumerable<IExtendedBase> ExtendedCollection
        {
            get { return SubjectScoreDetails == null ? new List<IExtendedBase>() : SubjectScoreDetails.Cast<IExtendedBase>(); }
        }



        /// <summary>
        /// 合计
        /// </summary>
        [DisplayName("合计")]
        public Int32 Total { get; set; }

        /// <summary>
        /// 总分平均
        /// </summary>
        [DisplayName("总分平均")]
        public Double TotalAverage { get; set; }
    }
}
