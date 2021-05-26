using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;


namespace STK.DataTable
{
    public class DataTableReader
    {
        public const char ARRAY_START = '[';
        public const char ARRAY_END = ']';
        public const char ARRAY_SEPARATOR = ',';
        public const char FIELD_SEPARATOR = ';';
        public const char DESCRIPTOR = '=';

        public static readonly char[] TRIMED_CHARACTERS = new char[] { ' ', '\n' };
        private static readonly Type LIST_TYPE;


        static DataTableReader()
        {
            LIST_TYPE = typeof(List<>);
        }


        public DataTable ReadExcelWorksheet(Worksheet worksheet, out Type rowType)
        {
            Range range = worksheet.UsedRange;
            int columnCount = range.Columns.Count;
            int rowCount = range.Rows.Count;


            rowType = null;
            dynamic dataTable = null;
            List<FieldInfo> columnInfos = null;

            for (int r = 1; r <= rowCount; ++r)
            {
                Range row = range.Rows[r];
                string rowDefinition = row.Cells[1, 1].Value?.Trim(TRIMED_CHARACTERS) ?? "";

                if (columnInfos == null)
                {
                    if (rowDefinition.StartsWith("###"))
                    {
                        switch (rowDefinition.Substring(3).Trim(TRIMED_CHARACTERS).ToUpper())
                        {
                            case "TABLE_TYPE":
                                Type dataTableType = GetType(row.Cells[1, 2].Value);
                                rowType = dataTableType.BaseType.GenericTypeArguments[0];
                                dataTable = Activator.CreateInstance(dataTableType, new object[] { worksheet.Name });
                                break;


                            case "FIELD_NAME":
                                if (rowType == null)
                                {
                                    throw new Exception();
                                }

                                columnInfos = new List<FieldInfo> { null, null };
                                for (int c = 2; c <= columnCount; ++c)
                                {
                                    columnInfos.Add(rowType.GetField(row.Cells[1, c].Value, BindingFlags.IgnoreCase | BindingFlags.NonPublic | BindingFlags.Instance));
                                }

                                break;
                        }
                    }
                }
                else if (string.IsNullOrEmpty(rowDefinition))
                {
                    dynamic dataTableRow = ReadRow(rowType, row, columnInfos, dataTable, new DataTableRow.Metadata(r));
                    dataTable.AddRow(dataTableRow);
                }
            }


            return dataTable;
        }

        private dynamic ReadRow(Type rowType, Range input, List<FieldInfo> columnInfos, DataTable dataTable, DataTableRow.Metadata metadata)
        {
            dynamic dataTableRow = Activator.CreateInstance(rowType, new object[] { dataTable, metadata });

            
            for (int c = 2; c < columnInfos.Count; ++c)
            {
                FieldInfo columnInfo = columnInfos[c];
                columnInfo.SetValue(dataTableRow, ReadString(columnInfo.FieldType, input.Cells[1, c].Value));
            }


            return dataTableRow;
        }

        public dynamic ReadString(Type type, string input)
        {
            input = input.Trim(TRIMED_CHARACTERS);


            if (type.IsArray)
            {
                return ReadArrayField(type, input.Trim(ARRAY_START, ARRAY_END));
            }


            return ReadField(type, input);
        }

        private dynamic ReadField(Type type, string input)
        {
            Type dataTableColumnType = type.GetInterface("IDataTableColumnType");
            if (dataTableColumnType != null)
            {
                dynamic obj = Activator.CreateInstance(type);
                dataTableColumnType.GetMethods()[0].Invoke(obj, new object[] { input, 0 });


                IEnumerable<MethodInfo> methodInfos = type.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic).Where(f => f.GetCustomAttributes(false).OfType<OnDeserializedAttribute>().Count() > 0);
                if (methodInfos.Count() > 0)
                {
                    methodInfos.First().Invoke(obj, new object[] { null });
                }


                return obj;
            }


            if (type.Equals(typeof(int)))
            {
                return int.Parse(input);
            }

            if (type.Equals(typeof(float)))
            {
                return float.Parse(input);
            }

            if (type.Equals(typeof(bool)))
            {
                return bool.Parse(input);
            }

            if (type.Equals(typeof(string)))
            {
                return input;
            }


            throw new Exception();
        }


        private dynamic ReadArrayField(Type type, string input)
        {
            Type elementType = type.GetElementType();


            if (elementType.IsArray)
            {
                dynamic list = Activator.CreateInstance(LIST_TYPE.MakeGenericType(elementType));
                for (int i = 0; i != -1 && i < input.Length; ++i)
                {
                    char c = input[i];
                    if (c != ' ' && c != '\n')
                    {
                        if (c == ARRAY_SEPARATOR)
                        {
                            list.Add(null);
                        }
                        else
                        {
                            if (c != ARRAY_START)
                            {
                                throw new Exception();
                            }

                            int j = input.IndexOf(ARRAY_END, i);
                            if (j == -1)
                            {
                                throw new Exception();
                            }

                            ++i;
                            list.Add(ReadArrayField(elementType, input.Substring(i, j - i)));


                            for (i = j + 1; input[i] != ARRAY_SEPARATOR; ++i) { }
                        }
                    }
                }

                return list.ToArray();
            }


            Type dataTableColumnType = elementType.GetInterface("IDataTableColumnType");
            if (dataTableColumnType != null)
            {
                MethodInfo generateFromSourceMethodInfo = dataTableColumnType.GetMethods()[0];

                dynamic list = Activator.CreateInstance(LIST_TYPE.MakeGenericType(elementType));
                for (int i = 0; i < input.Length; ++i)
                {
                    char c = input[i];
                    if (c != ' ' && c != '\n')
                    {
                        if (c == ARRAY_SEPARATOR)
                        {
                            list.Add(null);
                        }
                        else
                        {
                            dynamic obj = Activator.CreateInstance(elementType);
                            list.Add(obj);

                            i = (int)generateFromSourceMethodInfo.Invoke(obj, new object[] { input, i }) + 1;
                        }
                    }
                }

                return list.ToArray();
            }


            string[] splitedInput = input.Split(ARRAY_SEPARATOR);
            if (elementType.Equals(typeof(string)))
            {
                return splitedInput;
            }


            int N = splitedInput.Length;
            object[] array = new object[N];
            for (int i = 0; i < N; ++i)
            {
                array[i] = ReadField(elementType, splitedInput[i]);
            }


            return null;
        }


        protected virtual Type GetType(string name) => Type.GetType(name, true, true);
    }
}
