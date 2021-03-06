using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;


namespace STK.DataTable
{
    public sealed class TextTableExport
    {
        public static readonly string SPLIT_STRING = "\n<<<NEXT>>>\n";
        public static readonly string TEXT_KEY_FILE = "KEYS.txt";


        public static void ExportExcelWorkbook(Workbook workbook, string directory)
        {
            Sheets worksheets = workbook.Worksheets;


            char lastCharacterInDirectory = directory[directory.Length - 1];
            if (lastCharacterInDirectory != '/' && lastCharacterInDirectory != '\\')
            {
                directory += '\\';
            }


            Worksheet firstWorkSheet = null;
            Range firstSheetRange = null;
            foreach (Worksheet worksheet in worksheets)
            {
                firstWorkSheet = worksheet;
                firstSheetRange = worksheet.UsedRange;
                break;
            }


            int columnCount = firstSheetRange.Columns.Count;
            List<string> languages = new List<string> { null, null, null };

            for (int c = 3; c <= columnCount; ++c)
            {
                string expectedLanguage = firstSheetRange.Cells[1, c].Value?.Trim(DataTableReader.TRIMED_CHARACTERS);

                if (string.IsNullOrEmpty(expectedLanguage))
                {
                    throw new Exception(string.Format("[TextTableExporter] Undefined language\n  Worksheet: {0}\n  Column: {1}", firstWorkSheet.Name, c));
                }


                foreach (Worksheet worksheet in worksheets)
                {
                    Range range = worksheet.UsedRange;

                    if (range.Columns.Count != columnCount)
                    {
                        throw new Exception(string.Format("[TextTableExport] Unmatched column count\n  Worksheet: {0}\n  Expected:{1}\n  Actual:{2}", worksheet.Name, columnCount, range.Columns.Count));
                    }
                    else
                    {
                        string language = range.Cells[1, c].Value?.Trim(DataTableReader.TRIMED_CHARACTERS);

                        if (string.IsNullOrEmpty(language))
                        {
                            throw new Exception(string.Format("[TextTableExport] Undefined language\n  Worksheet: {0}\n  Column: {1}\n  Expected:{2}", worksheet.Name, c, expectedLanguage));
                        }
                        else if (language != expectedLanguage)
                        {
                            throw new Exception(string.Format("[TextTableExport] Unmatched language\n  Worksheet: {0}\n  Column: {1}\n  Expected:{2}\n  Actual:{3}", worksheet.Name, c, expectedLanguage, range.Cells[1, c].Value.Trim(' ', '\n')));
                        }
                    }
                }


                languages.Add(expectedLanguage);
            }


            List<List<bool>> doExports = new List<List<bool>>();
            HashSet<string> keys = new HashSet<string>();


            string contents = "";

            foreach (Worksheet worksheet in worksheets)
            {
                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                if (rowCount < 2)
                {
                    doExports.Add(null);
                    continue;
                }


                List<bool> doExport = new List<bool>() { false, false };

                for (int r = 2; r <= rowCount; ++r)
                {
                    if (!string.IsNullOrEmpty(range.Cells[r, 1].Value?.Trim(DataTableReader.TRIMED_CHARACTERS)))
                    {
                        doExport.Add(false);
                    }
                    else
                    {
                        string key = range.Cells[r, 2].Value?.Trim(DataTableReader.TRIMED_CHARACTERS);

                        if (string.IsNullOrEmpty(key))
                        {
                            doExport.Add(false);
                        }
                        else if (keys.Contains(key))
                        {
                            throw new Exception(string.Format("[TextTableExport] Duplicated key\n  Worksheet: {0}\n  Row: {1}\n  Key: {2}", worksheet.Name, r, key));
                        }
                        else
                        {
                            doExport.Add(true);
                            contents += key +'\n';
                        }
                    }
                }


                doExports.Add(doExport);
            }


            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(directory + TEXT_KEY_FILE, contents.Substring(0, contents.Length - 1));


            int i = 0;
            for (int c = 3; c <= columnCount; ++c)
            {
                contents = "";
                foreach (Worksheet worksheet in worksheets)
                {
                    Range range = worksheet.UsedRange;

                    int rowCount = range.Rows.Count;
                    if (rowCount < 2)
                    {
                        continue;
                    }


                    List<bool> doExport = doExports[i++];

                    for (int r = 2; r <= rowCount; ++r)
                    {
                        if (doExport[r])
                        {
                            contents += range.Cells[r, c].Value + SPLIT_STRING;
                        }
                    }
                }

                File.WriteAllText(directory + languages[c] + ".txt", contents.Substring(0, contents.Length - SPLIT_STRING.Length));
            }
        }


        public static void ExportExcelWorksheet(Worksheet worksheet, string directory)
        {
            Range range = worksheet.UsedRange;

            int rowCount = range.Rows.Count;
            if (rowCount < 2)
            {
                return;
            }


            List<bool> doExport = new List<bool>();
            HashSet<string> keys = new HashSet<string>();


            int r = 2;
            while (!string.IsNullOrEmpty(range.Cells[r, 1].Value?.Trim(DataTableReader.TRIMED_CHARACTERS)))
            {
                ++r;
            }
            int firstRowIndex = r;
            doExport.Add(true);

            string contents = range.Cells[r++, 2].Value.Trim(DataTableReader.TRIMED_CHARACTERS);
            for (; r <= rowCount; ++r)
            {
                if (!string.IsNullOrEmpty(range.Cells[r, 1].Value?.Trim(DataTableReader.TRIMED_CHARACTERS)))
                {
                    doExport.Add(false);
                }
                else
                {
                    string key = range.Cells[r, 2].Value?.Trim(DataTableReader.TRIMED_CHARACTERS);

                    if (string.IsNullOrEmpty(key))
                    {
                        doExport.Add(false);
                    }
                    else if (keys.Contains(key))
                    {
                        throw new Exception("Duplicated key: " + key);
                    }
                    else
                    {
                        doExport.Add(true);
                        contents += '\n' + key;
                    }
                }
            }


            char lastCharacterInDirectory = directory[directory.Length - 1];
            if (lastCharacterInDirectory != '/' && lastCharacterInDirectory != '\\')
            {
                directory += '\\';
            }

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(directory + TEXT_KEY_FILE, contents);


            int columnCount = range.Columns.Count;
            for (int c = 3; c <= columnCount; ++c)
            {
                string language = range.Cells[1, c].Value.Trim(DataTableReader.TRIMED_CHARACTERS);

                if (language == "KEYS")
                {
                    throw new Exception();
                }


                r = firstRowIndex;
                contents = range.Cells[r++, c].Value;

                for (; r <= rowCount; ++r)
                {
                    if (doExport[r - firstRowIndex])
                    {
                        contents += SPLIT_STRING + range.Cells[r, c].Value;
                    }
                }


                File.WriteAllText(directory + language + ".txt", contents);
            }
        }
    }
}
