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
            Type customExcelReadingInterface = null;


            int r = 1;

            for (; r <= rowCount; ++r)
            {
                Range row = range.Rows[r];
                string rowDefinition = row.Cells[1, 1].MergeArea.Cells[1,1].Value?.Trim(TRIMED_CHARACTERS) ?? "";

                if (rowDefinition.StartsWith("###") && rowDefinition.Substring(3).TrimStart(TRIMED_CHARACTERS).ToUpper() == "TABLE_TYPE")
                {
                    Type dataTableType = GetType(row.Cells[1, 2].MergeArea.Cells[1,1].Value);
                    rowType = dataTableType.BaseType.GenericTypeArguments[0];
                    dataTable = Activator.CreateInstance(dataTableType, new object[] { worksheet.Name });
                    customExcelReadingInterface = rowType.GetInterface("ICustomExcelReading");
                    break;
                }
            }


            bool hasColumnInfos = false;

            for (; r <= rowCount; ++r)
            {
                Range row = range.Rows[r];
                string rowDefinition = row.Cells[1, 1].MergeArea.Cells[1,1].Value?.Trim(TRIMED_CHARACTERS) ?? "";

                if (rowDefinition.StartsWith("###") && rowDefinition.Substring(3).TrimStart(TRIMED_CHARACTERS).ToUpper() == "FIELD_NAME")
                {
                    hasColumnInfos = true;
                    break;
                }
            }

            if (!hasColumnInfos)
            {
                throw new Exception();
            }


            if (customExcelReadingInterface == null)
            {
                Range row = range.Rows[r];

                List<FieldInfo> columnInfos = new List<FieldInfo> { null, null };
                for (int c = 2; c <= columnCount; ++c)
                {
                    columnInfos.Add(rowType.GetField(row.Cells[1, c].MergeArea.Cells[1,1].Value.Trim(TRIMED_CHARACTERS), BindingFlags.IgnoreCase | BindingFlags.NonPublic | BindingFlags.Instance));
                }


                for (; r <= rowCount; ++r)
                {
                    row = range.Rows[r];

                    if (string.IsNullOrEmpty(row.Cells[1, 1].MergeArea.Cells[1,1].Value))
                    {
                        dataTable.AddRow(ReadExcelRow(rowType, row, columnInfos, dataTable, new DataTableRow.Metadata(r)));
                    }
                }
            }
            else
            {
                MethodInfo readRowMethodInfo = customExcelReadingInterface.GetMethods()[0];


                Dictionary<string, object> arg = new Dictionary<string, object>();
                Range row = range.Rows[r];

                List<string> columnInfos = new List<string> { null, null };
                for (int c = 2; c <= columnCount; ++c)
                {
                    string columnInfo = row.Cells[1, c].MergeArea.Cells[1,1].Value.Trim(TRIMED_CHARACTERS);
                    columnInfos.Add(columnInfo);
                    arg.Add(columnInfo, "");
                }


                for (; r <= rowCount; ++r)
                {
                    row = range.Rows[r];

                    if (string.IsNullOrEmpty(row.Cells[1, 1].MergeArea.Cells[1,1].Value))
                    {
                        dynamic dataTableRow = Activator.CreateInstance(rowType, new object[] { dataTable, new DataTableRow.Metadata(r) });


                        for (int c = 2; c <= columnCount; ++c)
                        {
                            arg[columnInfos[c]] = row.Cells[1, c].MergeArea.Cells[1,1].Value;
                        }
                        
                        readRowMethodInfo.Invoke(dataTableRow, new object[] { arg });


                        dataTable.AddRow(dataTableRow);
                    }
                }
            }


            return dataTable;
        }


        private dynamic ReadExcelRow(Type rowType, Range input, List<FieldInfo> columnInfos, DataTable dataTable, DataTableRow.Metadata metadata)
        {
            dynamic dataTableRow = Activator.CreateInstance(rowType, new object[] { dataTable, metadata });


            for (int c = 2; c < columnInfos.Count; ++c)
            {
                FieldInfo columnInfo = columnInfos[c];
                Type columnType = columnInfo.FieldType;
                dynamic value = input.Cells[1, c].MergeArea.Cells[1,1].Value;

                if (value is string)
                {
                    columnInfo.SetValue(dataTableRow, ReadString(columnType, value));
                }
                else
                {
                    if (value.GetType() == columnType)
                    {
                        columnInfo.SetValue(dataTableRow, value);
                    }
                    else if (value is double)
                    {
                        if (columnType == typeof(int))
                        {
                            columnInfo.SetValue(dataTableRow, (int)(double)value);
                        }
                        else if (columnType == typeof(float))
                        {
                            columnInfo.SetValue(dataTableRow, (float)(double)value);
                        }
                        else
                        {
                            throw new Exception();
                        }
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
            }


            return dataTableRow;
        }

        public dynamic ReadString(Type type, string input)
        {
            input = input.Trim(TRIMED_CHARACTERS);


            if (type.IsArray)
            {
                return ReadArray(type, input.Trim(ARRAY_START, ARRAY_END));
            }


            return ReadVariable(type, input.Trim(ARRAY_START, ARRAY_END));
        }


        protected virtual Type GetType(string name) => Type.GetType(name, true, true);


        private dynamic ReadVariable(Type type, string input)
        {
            Type dataTableColumnInterface = type.GetInterface("IDataTableColumnType");
            if (dataTableColumnInterface != null)
            {
                dynamic obj = Activator.CreateInstance(type);
                dataTableColumnInterface.GetMethods()[0].Invoke(obj, new object[] { input, 0 });


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

        private dynamic ReadArray(Type type, string input)
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
                            list.Add(ReadArray(elementType, input.Substring(i, j - i)));


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
                array[i] = ReadVariable(elementType, splitedInput[i]);
            }


            return null;
        }
    }
}
