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
        private class ColumnInfo
        {
            public readonly FieldInfo fieldInfo;
            public readonly Type type;
            public readonly int startIndex;
            public readonly int width;


            public ColumnInfo(Type rowType, Range row, int column)
            {
                Range mergeArea = row.Cells[1, column].MergeArea;
                dynamic v = mergeArea.Cells[1, 1].Value;

                fieldInfo = rowType.GetField(mergeArea.Cells[1, 1].Value.Trim(TRIMED_CHARACTERS), BindingFlags.IgnoreCase | BindingFlags.NonPublic | BindingFlags.Instance);
                type = fieldInfo.FieldType;
                startIndex = column;
                width = mergeArea.Columns.Count;
            }
        }


        public const char ARRAY_START = '[';
        public const char ARRAY_END = ']';
        public const char ARRAY_SEPARATOR = ',';
        public const char FIELD_SEPARATOR = ';';
        public const char DESCRIPTOR = '=';

        public static readonly char[] TRIMED_CHARACTERS = new char[] { ' ', '\n' };
        private static readonly Type LIST_TYPE;
        private static readonly string ICUSTOMEXCELROWREADING_INTERFACE;
        private static readonly string IDATATABLECOLUMNTYPE_INTERFACE;
        private static readonly MethodInfo ICUSTOMEXCELROWREADING_GENERATEFROMSOURCE;
        private static readonly MethodInfo IDATATABLECOLUMNTYPE_GENERATEFROMSOURCE;


        static DataTableReader()
        {
            LIST_TYPE = typeof(List<>);
            ICUSTOMEXCELROWREADING_INTERFACE = typeof(ICustomExcelRowReading).Name;
            IDATATABLECOLUMNTYPE_INTERFACE = typeof(IDataTableColumnType).Name;
            ICUSTOMEXCELROWREADING_GENERATEFROMSOURCE = typeof(ICustomExcelRowReading).GetMethods()[0];
            IDATATABLECOLUMNTYPE_GENERATEFROMSOURCE = typeof(IDataTableColumnType).GetMethods()[0];
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


            bool hasTableType = false;

            for (; r <= rowCount; ++r)
            {
                Range row = range.Rows[r];
                string rowDefinition = GetExcelCellValue(row, 1, 1)?.Trim(TRIMED_CHARACTERS) ?? "";

                if (rowDefinition.StartsWith("###") && rowDefinition.Substring(3).TrimStart(TRIMED_CHARACTERS).ToUpper() == "TABLE_TYPE")
                {
                    Type dataTableType = GetType(GetExcelCellValue(row, 1, 2));
                    rowType = dataTableType.BaseType.GenericTypeArguments[0];
                    dataTable = Activator.CreateInstance(dataTableType, new object[] { worksheet.Name });
                    customExcelReadingInterface = rowType.GetInterface(ICUSTOMEXCELROWREADING_INTERFACE);
                    hasTableType = true;
                    break;
                }
            }

            if (!hasTableType)
                return null;


            bool hasColumnInfos = false;

            for (; r <= rowCount; ++r)
            {
                string rowDefinition = GetExcelCellValue(range.Rows[r], 1, 1)?.Trim(TRIMED_CHARACTERS) ?? "";

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

                List<ColumnInfo> columnInfos = new List<ColumnInfo>();
                for (int c = 2; c <= columnCount; ++c)
                {
                    ColumnInfo columnInfo = new ColumnInfo(rowType, row, c);
                    columnInfos.Add(columnInfo);

                    c += columnInfo.width - 1;
                }


                for (; r <= rowCount; ++r)
                {
                    row = range.Rows[r];

                    if (string.IsNullOrEmpty(GetExcelCellValue(row, 1, 1)))
                    {
                        dataTable.AddRow(ReadExcelRow(rowType, row, columnInfos, dataTable, new DataTableRow.Metadata(r)));
                    }
                }
            }
            else
            {
                Dictionary<string, object> arg = new Dictionary<string, object>();
                Range row = range.Rows[r];

                List<string> columnInfos = new List<string> { null, null };
                for (int c = 2; c <= columnCount; ++c)
                {
                    string columnInfo = GetExcelCellValue(row, 1, c).Trim(TRIMED_CHARACTERS);
                    columnInfos.Add(columnInfo);
                    arg.Add(columnInfo, "");
                }


                for (; r <= rowCount; ++r)
                {
                    row = range.Rows[r];

                    if (string.IsNullOrEmpty(GetExcelCellValue(row, 1, 1)))
                    {
                        dynamic dataTableRow = Activator.CreateInstance(rowType, new object[] { dataTable, new DataTableRow.Metadata(r) });


                        for (int c = 2; c <= columnCount; ++c)
                        {
                            arg[columnInfos[c]] = GetExcelCellValue(row, 1, c);
                        }

                        ICUSTOMEXCELROWREADING_GENERATEFROMSOURCE.Invoke(dataTableRow, new object[] { arg });


                        dataTable.AddRow(dataTableRow);
                    }
                }
            }


            return dataTable;
        }


        public dynamic Parse(Type type, string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return null;
            }


            input = input.Trim(TRIMED_CHARACTERS);


            if (type.IsArray)
            {
                return ParseArray(type, input.Trim(ARRAY_START, ARRAY_END));
            }


            return ParseObject(type, input);
        }

        public dynamic ParseObject(Type type, string input)
        {
            Type dataTableColumnInterface = type.GetInterface(IDATATABLECOLUMNTYPE_INTERFACE);
            if (dataTableColumnInterface != null)
            {
                dynamic obj = Activator.CreateInstance(type);
                IDATATABLECOLUMNTYPE_GENERATEFROMSOURCE.Invoke(obj, new object[] { input, -1 });

                return obj;
            }


            if (type.IsEnum)
            {
                return Enum.Parse(type, input, true);
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
                return input.ToString();
            }


            throw new Exception();
        }

        public dynamic ParseArray(Type type, string input)
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
                            list.Add(ParseArray(elementType, input.Substring(i, j - i)));


                            for (i = j + 1; input[i] != ARRAY_SEPARATOR; ++i) { }
                        }
                    }
                }

                return list.ToArray();
            }


            Type dataTableColumnInterface = elementType.GetInterface(IDATATABLECOLUMNTYPE_INTERFACE);
            if (dataTableColumnInterface != null)
            {
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

                            i = (int)IDATATABLECOLUMNTYPE_GENERATEFROMSOURCE.Invoke(obj, new object[] { input, i }) + 1;
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
                array[i] = ParseObject(elementType, splitedInput[i]);
            }


            return null;
        }


        protected virtual Type GetType(string name) => Type.GetType(name, true, true);

        protected virtual dynamic GetExcelCellValue(Range range, int row, int column) => range.Cells[row, column].MergeArea.Cells[1, 1].Value;


        private dynamic ReadExcelRow(Type rowType, Range input, List<ColumnInfo> columnInfos, DataTable dataTable, DataTableRow.Metadata metadata)
        {
            dynamic dataTableRow = Activator.CreateInstance(rowType, new object[] { dataTable, metadata });


            foreach (ColumnInfo columnInfo in columnInfos)
            {
                FieldInfo fieldInfo = columnInfo.fieldInfo;
                Type type = columnInfo.type;


                int c = columnInfo.startIndex;

                if (columnInfo.width == 1)
                {
                    dynamic cellValue = GetExcelCellValue(input, 1, c);

                    if (cellValue is string)
                    {
                        fieldInfo.SetValue(dataTableRow, Parse(type, cellValue));
                    }
                    else
                    {
                        if (cellValue.GetType() == type)
                        {
                            fieldInfo.SetValue(dataTableRow, cellValue);
                        }
                        else if (cellValue is double)
                        {
                            if (type == typeof(int))
                            {
                                fieldInfo.SetValue(dataTableRow, (int)(double)cellValue);
                            }
                            else if (type == typeof(float))
                            {
                                fieldInfo.SetValue(dataTableRow, (float)(double)cellValue);
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
                else
                {
                    if (!type.IsArray)
                    {
                        throw new Exception();
                    }


                    Type elementType = type.GetElementType();
                    dynamic list = Activator.CreateInstance(LIST_TYPE.MakeGenericType(elementType));

                    int C = c + columnInfo.width;

                    for (; c < C; ++c)
                    {
                        dynamic cellValue = GetExcelCellValue(input, 1, c);

                        if (string.IsNullOrEmpty(cellValue))
                        {
                            list.Add(null);
                        }
                        else if (cellValue is string)
                        {
                            list.Add(Parse(elementType, cellValue));
                        }
                        else
                        {
                            if (cellValue.GetType() == elementType)
                            {
                                list.Add(cellValue);
                            }
                            else if (cellValue is double)
                            {
                                if (elementType == typeof(int))
                                {
                                    list.Add((int)(double)cellValue);
                                }
                                else if (elementType == typeof(float))
                                {
                                    list.Add((float)(double)cellValue);
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

                    dynamic value = list.ToArray();
                    fieldInfo.SetValue(dataTableRow, list.ToArray());
                }
            }


            return dataTableRow;
        }
    }
}
