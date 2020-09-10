using CT;
using CT.Data;
using System;

namespace AverageSpreadsExcelReport
{
    [Table]
    public struct LiveQuote
    {
        /// <summary>
        /// ID of the instance
        /// </summary>
        [Field(Flags = FieldFlags.ID)]
        public long ID;

        //[Field(Flags = FieldFlags.ID | FieldFlags.AutoIncrement, Length = 11)]
        //public int ID;

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
        //}

        //[Field(Name = "TimeStamp")]
        //public string TimeStampTemp;

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
        /// Bid price
        /// </summary>
        [Field]
        public double Bid;

        /// <summary>
        /// Ask price
        /// </summary>
        [Field]
        public double Ask;

        /// <summary>
        /// Spread
        /// </summary>
        [Field]
        public double Spread;

        /// <summary>
        /// Average spread
        /// </summary>
        [Field]
        public double SpreadAvg;
    }
}