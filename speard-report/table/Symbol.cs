

using CT;
using CT.Data;

namespace AverageSpreadsExcelReport
{
    /// <summary>
    /// Struct for Symbols Table in GBE Prime RW database
    /// </summary>
    [Table]
    public struct Symbol
    {
        /// <summary>
        /// ID of the instance
        /// </summary>
        [Field(Flags = FieldFlags.ID | FieldFlags.AutoIncrement)]
        public long ID;

        [Field]
        public string currencypairname;

        [Field]
        public string requestId;

        [Field]
        public int Digit;

        [Field]
        public bool LiveQuotes;

        public string SymbolName
        {
            get
            {
                return !string.IsNullOrEmpty(currencypairname) ? currencypairname.Replace("/", "") : null;
            }
        }
    }
}