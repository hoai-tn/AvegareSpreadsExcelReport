using CT;
using CT.Data;
using System;

namespace AverageSpreadsExcelReport
{
    /// <summary>
    /// Provides an average spread struture to hold data at specified time.
    /// </summary>
    [Table]
    public struct AverageSpread
    {
        /// <summary>
        /// ID of the instance
        /// </summary>
        [Field(Flags = FieldFlags.ID | FieldFlags.AutoIncrement)]
        public long ID;

        /// <summary>
        /// Time stamp
        /// </summary>
        [Field(Flags = FieldFlags.Index)]
        [DateTimeFormat(DateTimeKind.Utc, DateTimeType.BigIntHumanReadable)]
        public DateTime TimeStamp;

        //public DateTime TimeStamp
        //{
        //    get
        //    {
        //        return Cave.Text.StringTools.ParseDateTime(TimeStampTemp);
        //    }
        //    set
        //    {
        //        TimeStampTemp = value.ToLongTimeString();
        //    }
        //}

        //[Field(Name = "TimeStamp")]
        //public string TimeStampTemp;

        /// <summary>
        /// Period of time we will calculate an average value from data
        /// </summary>
        [Field]
        public int Duration;

        /// <summary>
        /// Broker name
        /// </summary>
        [Field]
        public string BrokerName;

        /// <summary>
        /// Symbol name
        /// </summary>
        [Field(Flags = FieldFlags.Index)]
        public string Symbol;

        /// <summary>
        /// Average spread
        /// </summary>
        [Field]
        public double AvgSpread;
    }
}