using Newtonsoft.Json;
using System.Collections.Generic;


namespace STK.DataTable
{
    [JsonObject]
    public abstract class DataTable<RowType> where RowType : DataTableRow
    {
        [JsonIgnore]
        public abstract IEnumerable<RowType> Rows { get; }


        internal abstract void AddRow(RowType row);
    }


    public abstract class ListDataTable<RowType> : DataTable<RowType> where RowType : DataTableRow
    {
        [JsonProperty] private List<RowType> rows;


        public RowType this[int index] => rows[index];


        public override IEnumerable<RowType> Rows => rows;


        internal override void AddRow(RowType row)
        {
            rows.Add(row);
        }
    }


    public abstract class DictionaryDataTable<RowType, KeyType> : DataTable<RowType> where RowType : DictionaryDataTableRow<KeyType>
    {
        [JsonProperty] private Dictionary<KeyType, RowType> rows = new Dictionary<KeyType, RowType>();


        public RowType this[KeyType key] => rows[key];


        public override IEnumerable<RowType> Rows => rows.Values;


        internal override void AddRow(RowType row)
        {
            rows.Add(row.Key, row);
        }
    }
}
