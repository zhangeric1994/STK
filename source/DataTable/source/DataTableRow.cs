namespace STK.DataTable
{
    public abstract class DataTableRow
    {
        public int Index { get; private set; }


        protected DataTableRow() { }
    }


    public abstract class DictionaryDataTableRow<T> : DataTableRow
    {
        public abstract T Key { get; }
    }
}
