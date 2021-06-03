using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;


namespace STK.DataTable
{
    public sealed class DataTableManager
    {
        public static readonly DataTableManager Instance = new DataTableManager();


        private List<DataTable> dataTableList = new List<DataTable>();
        private Dictionary<string, DataTable> dataTableDictionary = new Dictionary<string, DataTable>();
        internal readonly JsonSerializer serializer = new JsonSerializer() { Formatting = Formatting.Indented, TypeNameHandling = TypeNameHandling.Objects };


        public int DataTableCount => dataTableList.Count;


        private DataTableManager() { }


        public T GetDataTableByUID<T>(int uid) where T : DataTable => dataTableList[uid - 1] as T;

        public T GetDataTableByName<T>(string name) where T : DataTable => dataTableDictionary.TryGetValue(name, out DataTable dataTable) ? dataTable as T : null;


        public T GetDataByUID<T>(int uid) where T : DataTableRow => GetDataTableByUID<DataTable<T>>(DataTableRow.GetTableUIDFromUID(uid)).GetDataByID<T>(DataTableRow.GetIDFromUID(uid));


        public void ImportFromJSON(string file)
        {
            if (!File.Exists(file))
            {
                return;
            }


            using (TextReader tr = File.OpenText(file))
            {
                using (JsonTextReader jtr = new JsonTextReader(tr))
                {
                    DataTable dataTable = serializer.Deserialize(jtr) as DataTable;
                    dataTable.UID = AddDataTable(dataTable);
                }
            }
        }

        public void ImportAllFromJSON(string directory)
        {
            foreach (string file in Directory.EnumerateFiles(directory, "*.json"))
            {
                ImportFromJSON(file);
            }
        }


        internal int AddDataTable(DataTable dataTable)
        {
            string key = dataTable.name;

            if (dataTableDictionary.TryGetValue(key, out DataTable oldDataTable))
            {
                dataTableDictionary[key] = dataTable;

                int uid = oldDataTable.UID;
                dataTableList[uid - 1] = dataTable;

                return uid;
            }


            dataTableDictionary.Add(key, dataTable);
            dataTableList.Add(dataTable);

            return DataTableCount;
        }
    }
}
