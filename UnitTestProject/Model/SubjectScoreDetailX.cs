using ExcelHelper.Operating;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestProject.Model
{
    public struct SubjectScoreDetailX : IExtendedBase
    {
        public object TypeValue
        {
            get
            {
                return this.Score;
            }
            set
            {
                Score = (Double?)value;
            }
        }

        public Double? Score { get; set; }

        public Int32 TypeId { get; set; }

        public String TypeName { get; set; }
    }
}
