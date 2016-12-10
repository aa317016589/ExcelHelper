using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel;
using ExcelHelper.Operating;


namespace UnitTestProject.Model
{
    /// <summary>
    /// 积分兑换券
    /// </summary>
    public class PointsCoupon
    {
        private Int32 _pointsCouponId;
        [Ignore]
        public Int32 PointsCouponId
        {
            get { return _pointsCouponId; }
            set { _pointsCouponId = value; }
        }

        private DateTime _createDateTime;

        /// <summary>
        /// 创造日期
        /// </summary>
        [DisplayName("创造日期")]
        [Ignore]
        public DateTime CreateDateTime
        {
            get { return _createDateTime; }
            set { _createDateTime = value; }
        }

        private String _redemptionCode;
        /// <summary>
        /// 兑换码
        /// </summary>
        [DisplayName("兑换码")]
        public String RedemptionCode
        {
            get { return _redemptionCode; }
            set { _redemptionCode = value; }
        }


        private Int32 parValue;
        /// <summary>
        /// 面值
        /// </summary>
        [DisplayName("面值")]
        public Int32 ParValue
        {
            get { return parValue; }
            set { parValue = value; }
        }

        private DateTime _pastDateTime;
        /// <summary>
        /// 过期日期
        /// </summary>
        [DisplayName("过期日期")]
        public DateTime PastDateTime
        {
            get { return _pastDateTime; }
            set { _pastDateTime = value; }
        }


        private PointsCouponTypes _pointsCouponTypes;
        /// <summary>
        /// 类型
        /// </summary>
        [DisplayName("类型")]
        public PointsCouponTypes PointsCouponTypes
        {
            get { return _pointsCouponTypes; }
            set { _pointsCouponTypes = value; }
        }


        private String userName;
        /// <summary>
        /// 使用的会员账号
        /// </summary>
        [DisplayName("账号")]
        public String UserName
        {
            get { return userName; }
            set { userName = value; }
        }

        private Boolean isUsed;
        [Ignore]
        public Boolean IsUsed
        {
            get { return isUsed; }
            set { isUsed = value; }
        }

        public PointsCoupon()
        {
            RedemptionCode = new Random().Next(1, 99999).ToString();

            _createDateTime = DateTime.Now;

            isUsed = false;
        }

    }

    /// <summary>
    /// 积分兑换券类型
    /// </summary>
    public enum PointsCouponTypes
    {
        [Description("电子")]
        Electronic = 1,
        [Description("纸质")]
        Paper
    }
}