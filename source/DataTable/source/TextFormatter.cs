using System;
using System.Collections.Generic;


namespace STK.DataTable
{
    public class TextFormatter : ITextFormatterNode, IDataTableColumnType
    {
        private enum ReadingType
        {
            String,
            Text,
            TextVariable,
        }


        private List<ITextFormatterNode> nodes;


        string ITextFormatterNode.GetText(Dictionary<string, string> dictionary)
        {
            string text = "";

            foreach (ITextFormatterNode node in nodes)
            {
                text += node.GetText(dictionary);
            }

            return text;
        }


        int IDataTableColumnType.GenerateFromSource(string input, int leftIndex)
        {
            if (leftIndex != -1)
            {
                throw new Exception();
            }


            nodes = new List<ITextFormatterNode>();

            
            leftIndex = 0;
            int rightIndex = 0;
            ReadingType currentReadingType = ReadingType.String;
            while (rightIndex < input.Length)
            {
                switch (input[rightIndex])
                {
                    case '[':
                        if (rightIndex != leftIndex)
                        {
                            nodes.Add(new StringNode(input.Substring(leftIndex, rightIndex - leftIndex)));
                        }

                        leftIndex = ++rightIndex;
                        currentReadingType = ReadingType.TextVariable;
                        break;


                    case ']':
                        if (currentReadingType != ReadingType.TextVariable)
                        {
                            throw new Exception();
                        }

                        nodes.Add(new TextVariableNode(input.Substring(leftIndex, rightIndex - leftIndex)));

                        leftIndex = ++rightIndex;
                        break;


                    case '{':
                        if (rightIndex != leftIndex)
                        {
                            nodes.Add(new StringNode(input.Substring(leftIndex, rightIndex - leftIndex)));
                        }

                        leftIndex = ++rightIndex;
                        currentReadingType = ReadingType.Text;
                        break;


                    case '}':
                        if (currentReadingType != ReadingType.Text)
                        {
                            throw new Exception();
                        }

                        nodes.Add(new TextNode(input.Substring(leftIndex, rightIndex - leftIndex)));

                        leftIndex = ++rightIndex;
                        break;


                    default:
                        ++rightIndex;
                        break;
                }
            }


            if (leftIndex != input.Length)
            {
                nodes.Add(new StringNode(input.Substring(leftIndex)));
            }


            return -1;
        }
    }


    public interface ITextFormatterNode
    {
        string GetText(Dictionary<string, string> dictionary);
    }


    public readonly struct StringNode : ITextFormatterNode
    {
        public readonly string text;


        public StringNode(string text)
        {
            this.text = text;
        }


        string ITextFormatterNode.GetText(Dictionary<string, string> dictionary)
        {
            return text;
        }
    }


    public readonly struct TextNode : ITextFormatterNode
    {
        public readonly string name;


        public TextNode(string name)
        {
            this.name = name;
        }


        string ITextFormatterNode.GetText(Dictionary<string, string> dictionary)
        {
            return TextManager.Instance.GetText(name);
        }
    }


    public readonly struct TextVariableNode : ITextFormatterNode
    {
        public readonly string name;


        public TextVariableNode(string name)
        {
            this.name = name;
        }


        string ITextFormatterNode.GetText(Dictionary<string, string> dictionary)
        {
            if (dictionary.TryGetValue(name, out string text))
            {
                return text;
            }

            return string.Format("[ {0}? ]", name);
        }
    }
}
