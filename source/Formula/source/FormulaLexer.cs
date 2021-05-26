using System.Collections.Generic;


namespace STK.Formula
{
    public class FormulaLexer
    {
        public const char VARIABLE_START = '[';
        public const char VARIABLE_END = ']';


        public List<FormulaToken> GenerateTokens(string input)
        {
            List<FormulaToken> result;
            if (GenerateTokens(input, out result))
            {
                return result;
            }

            return null;
        }


        public bool GenerateTokens(string input, out List<FormulaToken> output)
        {
            output = new List<FormulaToken>();


            for (int i = 0; i < input.Length;)
            {
                char c = input[i];

                if (!char.IsWhiteSpace(c))
                {
                    if (char.IsDigit(c))
                    {
                        int numDots = 0;


                        int count = 1;
                        while (i + count < input.Length)
                        {
                            c = input[i + count];

                            if (char.IsDigit(c))
                            {
                                ++count;
                            }
                            else if (c == '.')
                            {
                                if (++numDots > 1)
                                {
                                    return false;
                                }

                                ++count;
                            }
                            else
                            {
                                break;
                            }
                        }


                        output.Add(new FormulaToken("CONSTANT", input.Substring(i, count)));


                        i += count;
                    }
                    else
                    {
                        switch (c)
                        {
                            case VARIABLE_START:
                                ++i;


                                int count = 0;
                                for (; i + count < input.Length; ++count)
                                {
                                    if (input[i + count] == VARIABLE_END)
                                    {
                                        if (count == 0)
                                        {
                                            return false;
                                        }

                                        break;
                                    }
                                }


                                output.Add(new FormulaToken("VARIABLE", input.Substring(i, count)));


                                i += count + 1;
                                break;


                            case '+':
                                output.Add(new FormulaToken("ADDITION"));
                                ++i;
                                break;


                            case '-':
                                output.Add(new FormulaToken("SUBTRACTION"));
                                ++i;
                                break;


                            case '*':
                                output.Add(new FormulaToken("MULTIPLICATION"));
                                ++i;
                                break;


                            case '/':
                                output.Add(new FormulaToken("DIVISION"));
                                ++i;
                                break;


                            case '(':
                                output.Add(new FormulaToken("LEFT_PARENTHESIS"));
                                ++i;
                                break;


                            case ')':
                                output.Add(new FormulaToken("RIGHT_PARENTHESIS"));
                                ++i;
                                break;


                            default:
                                if (!HandleUnexpectedCharacter(input, ref output, ref i))
                                {
                                    return false;
                                }
                                break;
                        }
                    }
                }
                else
                {
                    ++i;
                }
            }


            return true;
        }


        protected virtual bool HandleUnexpectedCharacter(string input, ref List<FormulaToken> output, ref int index)
        {
            return false;
        }
    }


    public class FormulaToken
    {
        public readonly string type;
        public readonly string text;


        public FormulaToken(string type, string text = "")
        {
            this.type = type.ToUpper();
            this.text = text;
        }


        public override string ToString()
        {
            return string.Format("[{0}: {1}]", type, text);
        }
    }
}
