namespace STK.DataTable
{
    public interface IDataTableColumnType
    {
        int GenerateFromSource(string input, int leftIndex = -1);
    }
}
