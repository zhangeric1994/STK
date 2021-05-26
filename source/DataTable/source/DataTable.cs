using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;


namespace STK.DataTable
{
    [JsonObject]
    public abstract class DataTable
    {
        public readonly string name;
        public readonly int uid;


        public abstract int Count { get; }


        public DataTable(string name)
        {
            this.name = name;
            this.uid = DataTableManager.Instance.AddDataTable(this);
        }


        public abstract T GetDataByID<T>(int id) where T : DataTableRow;


        public void ExportToJSON(string directory)
        {
            using (StreamWriter sw = File.CreateText(string.Format("{0}\\{1}.json", directory, name)))
            {
                using (JsonTextWriter jtw = new JsonTextWriter(sw))
                {
                    DataTableManager.Instance.serializer.Serialize(jtw, this);
                }
            }
        }
    }


    [JsonObject]
    public abstract class DataTable<RowType> : DataTable where RowType : DataTableRow
    {
        [JsonIgnore]
        public abstract IEnumerable<RowType> Rows { get; }


        public DataTable(string name) : base(name) { }


        internal abstract void AddRow(RowType row);
    }


    [JsonObject]
    public abstract class ListDataTable<RowType> : DataTable<RowType> where RowType : DataTableRow
    {
        [JsonProperty]
        protected List<RowType> rows = new List<RowType>();


        [JsonIgnore]
        public RowType this[int index] => rows[index];
        [JsonIgnore]
        public override int Count => rows.Count;
        [JsonIgnore]
        public override IEnumerable<RowType> Rows => rows;


        public ListDataTable(string name) : base(name) { }


        public override T GetDataByID<T>(int id) => rows[id - 1] as T;


        internal override void AddRow(RowType row)
        {
            rows.Add(row);
        }
    }


    [JsonObject]
    public abstract class DictionaryDataTable<RowType, KeyType> : ListDataTable<RowType> where RowType : DictionaryDataTableRow<KeyType>
    {
        [JsonIgnore]
        protected Dictionary<KeyType, RowType> rowDictionary = new Dictionary<KeyType, RowType>();


        [JsonIgnore]
        public RowType this[KeyType key] => rowDictionary[key];


        public DictionaryDataTable(string name) : base(name) { }


        [OnDeserialized]
        private void OnDeserialized(StreamingContext context)
        {
            foreach (RowType row in rows)
            {
                rowDictionary.Add(row.Key, row);
            }
        }
    }
}
