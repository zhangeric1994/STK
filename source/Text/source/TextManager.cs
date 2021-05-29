using System;
using System.IO;
using System.Collections.Generic;


namespace STK.Text
{
    public sealed class Text
    {
        internal string buffer = null;
        internal int startIndex = -1;
        internal int length = -1;
        internal Text next = null;


        public string Key { get; private set; }


        internal Text(string key)
        {
            Key = key;
        }


        public static implicit operator string(Text text) => text.ToString();


        public override string ToString() => buffer.Substring(startIndex, length);
    }


    public sealed class TextManager
    {
        public static readonly TextManager Instance = new TextManager();


        static TextManager() { }


        private string buffer = null;
        private Text textHead = null;
        private int numText = 0;
        private Dictionary<string, Text> dictionary = new Dictionary<string, Text>();


        public string CurrentLanguage { get; private set; } = "";


        public string GetText(string name)
        {
            if (dictionary.TryGetValue(name, out Text text))
            {
                return text;
            }

            return "{ " + name + "? }";
        }


        public void ImportTextKeys(string directory)
        {
            if (textHead != null)
            {
                throw new Exception();
            }


            char lastCharacterInDirectory = directory[directory.Length - 1];
            if (lastCharacterInDirectory != '/' && lastCharacterInDirectory != '\\')
            {
                directory += '\\';
            }


            string file = directory + TextTableExport.TEXT_KEY_FILE;


            if (File.Exists(file))
            {
                string[] keys = File.ReadAllText(file).Split('\n');

                numText = keys.Length;


                textHead = new Text("");
                Text previous = textHead;

                foreach (string key in keys)
                {
                    Text text = new Text(key);
                    previous.next = text;

                    dictionary.Add(key, text);

                    previous = text;
                }


                textHead = textHead.next;
            }
            else
            {
                throw new Exception();
            }
        }


        public void ImportText(string directory, string language)
        {
            if (textHead == null)
            {
                ImportTextKeys(directory);
            }


            if (language != CurrentLanguage)
            {
                char lastCharacterInDirectory = directory[directory.Length - 1];
                if (lastCharacterInDirectory != '/' && lastCharacterInDirectory != '\\')
                {
                    directory += '\\';
                }

                string file = directory + language + ".txt";


                if (File.Exists(file))
                {
                    string[] ss = File.ReadAllText(file).Split(new string[] { TextTableExport.SPLIT_STRING }, StringSplitOptions.None);
                    if (ss.Length != numText)
                    {
                        throw new Exception();
                    }


                    buffer = "";

                    Text current = textHead;
                    foreach (string s in ss)
                    {
                        current.startIndex = buffer.Length;
                        current.length = s.Length;

                        buffer += s;

                        current = current.next;
                    }


                    CurrentLanguage = language;


                    
                    for (current = textHead; current != null; current = current.next)
                    {
                        current.buffer = buffer;
                    }
                }
                else
                {
                    throw new Exception();
                }
            }
        }
    }
}
