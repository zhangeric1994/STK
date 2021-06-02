using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;


namespace STK.DataTable
{
    public sealed class TextTableExport
    {
        public static readonly string SPLIT_STRING = "\n[NEW]\n";
        public static readonly string TEXT_KEY_FILE = "KEYS.txt";


        public static void ExportExcelWorkbook(Workbook workbook, string directory)
        {
            Sheets worksheets = workbook.Worksheets;


            char lastCharacterInDirectory = directory[directory.Length - 1];
            if (lastCharacterInDirectory != '/' && lastCharacterInDirectory != '\\')
            {
                directory += '\\';
            }


            Range firstSheetRange = null;
            foreach (Worksheet worksheet in worksheets)
            {
                firstSheetRange = worksheet.UsedRange;
                break;
            }


            int columnCount = firstSheetRange.Columns.Count;
            List<string> languages = new List<string> { null, null, null };

            for (int c = 3; c <= columnCount; ++c)
            {
                string language = firstSheetRange.Cells[1, c].Value.Trim(' ', '\n');
                foreach (Worksheet worksheet in worksheets)
                {
                    Range range = worksheet.UsedRange;

                    if (range.Columns.Count != columnCount)
                    {
                        throw new Exception(string.Format("Unmatch column count: {0}\nExpected:{1} Actual:{2}", worksheet.Name, columnCount, range.Columns.Count));
                    }
                    else if (range.Cells[1, c].Value.Trim(' ', '\n') != language)
                    {
                        throw new Exception(string.Format("Unmatched language: {0}\nExpected:{1} Actual:{2}", worksheet.Name, language, range.Cells[1, c].Value.Trim(' ', '\n')));
                    }
                }

                languages.Add(language);
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
                    if (!string.IsNullOrEmpty(range.Cells[r, 1].Value?.Trim(' ', '\n')))
                    {
                        doExport.Add(false);
                    }
                    else
                    {
                        string key = range.Cells[r, 2].Value.Trim(' ', '\n');

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
                            contents += range.Cells[r, c].Value.Trim(' ') + SPLIT_STRING;
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
            while (range.Cells[r, 1].Value != "")
            {
                ++r;
            }
            int firstRowIndex = r;
            doExport.Add(true);

            string contents = range.Cells[r++, 2].Value.Trim(' ', '\n');
            for (; r <= rowCount; ++r)
            {
                if (!string.IsNullOrEmpty(range.Cells[r, 1].Value?.Trim(' ', '\n')))
                {
                    doExport.Add(false);
                }
                else
                {
                    string key = range.Cells[r, 2].Value.Trim(' ', '\n');

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
                string language = range.Cells[1, c].Value.Trim(' ', '\n');

                if (language == "KEYS")
                {
                    throw new Exception();
                }


                r = firstRowIndex;
                contents = range.Cells[r++, c].Value.Trim(' ');
                for (; r <= rowCount; ++r)
                {
                    if (doExport[r - firstRowIndex])
                    {
                        contents += SPLIT_STRING + range.Cells[r, c].Value.Trim(' ');
                    }
                }


                File.WriteAllText(directory + language + ".txt", contents);
            }
        }
    }
}
