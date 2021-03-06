using Newtonsoft.Json;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;


namespace STK.DataTable
{
    [JsonObject]
    public abstract class DataTable
    {
        public readonly string name;
        
        protected int uid;


        [JsonIgnore]
        public virtual int UID { get; internal set; }
        [JsonIgnore]
        public abstract int Count { get; }


        public DataTable(string name)
        {
            this.name = name;
        }


        public abstract T GetDataByID<T>(int id) where T : DataTableRow;


        public string ExportToJSON(string directory)
        {
            string file = string.Format("{0}\\{1}.json", directory, name);

            using (StreamWriter sw = File.CreateText(file))
            {
                using (JsonTextWriter jtw = new JsonTextWriter(sw))
                {
                    DataTableManager.Instance.serializer.Serialize(jtw, this);
                }
            }

            return file;
        }
    }


    [JsonObject]
    public abstract class DataTable<RowType> : DataTable, IEnumerable<RowType> where RowType : DataTableRow
    {
        [JsonProperty]
        protected List<RowType> rows = new List<RowType>();


        [JsonIgnore]
        public RowType this[int index] => rows[index];


        [JsonIgnore]
        public override int UID
        {
            internal set
            {
                if (value != uid)
                {
                    uid = value;

                    int id = 1;
                    foreach (RowType row in rows)
                    {
                        row.UID = uid << 28 + id++;
                    }
                }
            }
        }

        [JsonIgnore]
        public override int Count => rows.Count;
        

        public DataTable(string name) : base(name) { }


        public override T GetDataByID<T>(int id) => rows[id - 1] as T;


        internal virtual void AddRow(RowType row)
        {
            rows.Add(row);
        }


        IEnumerator<RowType> IEnumerable<RowType>.GetEnumerator() => rows.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => rows.GetEnumerator();
    }


    [JsonObject]
    public abstract class DictionaryDataTable<RowType, KeyType> : DataTable<RowType> where RowType : DictionaryDataTableRow<KeyType>
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


    [JsonObject]
    public abstract class DictionaryDataTable<RowType, KeyType1, KeyType2> : DataTable<RowType> where RowType : DictionaryDataTableRow<KeyType1, KeyType2>
    {
        [JsonIgnore]
        protected Dictionary<KeyType1, Dictionary<KeyType2, RowType>> rowDictionary = new Dictionary<KeyType1, Dictionary<KeyType2, RowType>>();


        [JsonIgnore]
        public RowType this[KeyType1 key1, KeyType2 key2] => rowDictionary[key1][key2];


        public DictionaryDataTable(string name) : base(name) { }


        [OnDeserialized]
        private void OnDeserialized(StreamingContext context)
        {
            foreach (RowType row in rows)
            {
                KeyType1 key1 = row.Key;

                if (!rowDictionary.ContainsKey(key1))
                {
                    rowDictionary.Add(key1, new Dictionary<KeyType2, RowType>() { { row.Key2, row } });
                }
                else
                {
                    rowDictionary[key1].Add(row.Key2, row);
                }
            }
        }
    }


    [JsonObject]
    public abstract class CatagorizedDataTable<RowType, KeyType> : DataTable<RowType> where RowType : DictionaryDataTableRow<KeyType>
    {
        [JsonIgnore]
        protected Dictionary<KeyType, List<RowType>> rowDictionary = new Dictionary<KeyType, List<RowType>>();


        [JsonIgnore]
        public List<RowType> this[KeyType key] => rowDictionary[key];


        public CatagorizedDataTable(string name) : base(name) { }


        [OnDeserialized]
        private void OnDeserialized(StreamingContext context)
        {
            foreach (RowType row in rows)
            {
                KeyType key = row.Key;

                if (!rowDictionary.ContainsKey(key))
                {
                    rowDictionary.Add(key, new List<RowType>() { { row } });
                }
                else
                {
                    rowDictionary[key].Add(row);
                }
            }
        }
    }


    [JsonObject]
    public abstract class CatagorizedDataTable<RowType, KeyType1, KeyType2> : DataTable<RowType> where RowType : DictionaryDataTableRow<KeyType1, KeyType2>
    {
        [JsonIgnore]
        protected Dictionary<KeyType1, Dictionary<KeyType2, List<RowType>>> rowDictionary = new Dictionary<KeyType1, Dictionary<KeyType2, List<RowType>>>();


        [JsonIgnore]
        public List<RowType> this[KeyType1 catagory1, KeyType2 catagory2] => rowDictionary[catagory1][catagory2];


        public CatagorizedDataTable(string name) : base(name) { }


        [OnDeserialized]
        private void OnDeserialized(StreamingContext context)
        {
            foreach (RowType row in rows)
            {
                KeyType1 key1 = row.Key;

                if (!rowDictionary.ContainsKey(key1))
                {
                    rowDictionary.Add(key1, new Dictionary<KeyType2, List<RowType>>() { { row.Key2, new List<RowType>() { row } } });
                }
                else
                {
                    Dictionary<KeyType2, List<RowType>> dictionary = rowDictionary[key1];
                    KeyType2 key2 = row.Key2;

                    if (!dictionary.ContainsKey(key2))
                    {
                        dictionary.Add(key2, new List<RowType>() { row });
                    }
                    else
                    {
                        dictionary[key2].Add(row);
                    }
                }
            }
        }
    }
}
