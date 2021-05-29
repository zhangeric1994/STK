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
        [JsonProperty]
        public readonly int uid;


        [JsonIgnore]
        public int ID { get => GetIDFromUID(uid); }
        [JsonIgnore]
        public int TableUID { get => GetTableUIDFromUID(uid); }


        public DataTableRow() { }


        public DataTableRow(DataTable dataTable, Metadata metadata)
        {
            this.metadata = metadata;

            if (dataTable != null)
            {
                this.uid = dataTable.uid << 24 + dataTable.Count + 1;
            }
        }
    }


    [JsonObject]
    public abstract class DictionaryDataTableRow<KeyType> : DataTableRow
    {
        [JsonIgnore]
        public abstract KeyType Key { get; }


        public DictionaryDataTableRow(DataTable dataTable, Metadata metadata) : base(dataTable, metadata) { }
    }


    [JsonObject]
    public abstract class DictionaryDataTableRow<KeyType1, KeyType2> : DataTableRow
    {
        [JsonIgnore]
        public abstract KeyType1 Key1 { get; }
        [JsonIgnore]
        public abstract KeyType2 Key2 { get; }


        public DictionaryDataTableRow(DataTable dataTable, Metadata metadata) : base(dataTable, metadata) { }
    }


    public interface ICustomExcelRowReading
    {
        void GenerateFromSource(Dictionary<string, object> input);
    }
}
