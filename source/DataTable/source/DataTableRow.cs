using Newtonsoft.Json;
using System.Collections.Generic;


namespace STK.DataTable
{
    [JsonObject]
    public abstract class DataTableRow
    {
        private const int FILTER_ID = 0x00111111;
        private const int FILTER_TABLEID = 0x11000000;


        public struct Metadata
        {
            public readonly int rowIndex;

            public Metadata(int rowIndex)
            {
                this.rowIndex = rowIndex;
            }
        }


        public static int GetIDFromUID(int uid) => uid & FILTER_ID;
        public static int GetTableUIDFromUID(int uid) => uid & FILTER_TABLEID;


        [JsonProperty]
        public readonly Metadata metadata;


        [JsonIgnore]
        public int UID { get; internal set; }
        [JsonIgnore]
        public int ID { get => GetIDFromUID(UID); }
        [JsonIgnore]
        public int TableUID { get => GetTableUIDFromUID(UID); }


        public DataTableRow() { }


        public DataTableRow(Metadata metadata)
        {
            this.metadata = metadata;
        }
    }


    [JsonObject]
    public abstract class DictionaryDataTableRow<KeyType> : DataTableRow
    {
        [JsonIgnore]
        public abstract KeyType Key { get; }


        public DictionaryDataTableRow(Metadata metadata) : base(metadata) { }
    }


    [JsonObject]
    public abstract class DictionaryDataTableRow<KeyType1, KeyType2> : DictionaryDataTableRow<KeyType1>
    {
        [JsonIgnore]
        public abstract KeyType2 Key2 { get; }


        public DictionaryDataTableRow(Metadata metadata) : base(metadata) { }
    }


    public interface ICustomExcelRowReading
    {
        void GenerateFromSource(Dictionary<string, object> input);
    }
}
