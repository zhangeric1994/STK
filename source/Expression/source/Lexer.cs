using System.Collections.Generic;


namespace STK.Expression
{
    public class Lexer
    {
        public const char VARIABLE_START = '[';
        public const char VARIABLE_END = ']';

        public static readonly char[] TRIMED_CHARACTERS = new char[] { ' ', '\n' };


        public List<Token> GenerateTokens(string input)
        {
            input = input.Trim(TRIMED_CHARACTERS);

            if (!string.IsNullOrEmpty(input) && GenerateTokens(input, out List<Token> tokens))
            {
                return tokens;
            }

            return null;
        }


        public bool GenerateTokens(string input, out List<Token> output)
        {
            output = new List<Token>();


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


                        output.Add(new Token("CONSTANT", input.Substring(i, count)));


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


                                output.Add(new Token("VARIABLE", input.Substring(i, count)));


                                i += count + 1;
                                break;


                            case '+':
                                output.Add(new Token("ADDITION"));
                                ++i;
                                break;


                            case '-':
                                output.Add(new Token("SUBTRACTION"));
                                ++i;
                                break;


                            case '*':
                                output.Add(new Token("MULTIPLICATION"));
                                ++i;
                                break;


                            case '/':
                                output.Add(new Token("DIVISION"));
                                ++i;
                                break;


                            case '(':
                                output.Add(new Token("LEFT_PARENTHESIS"));
                                ++i;
                                break;


                            case ')':
                                output.Add(new Token("RIGHT_PARENTHESIS"));
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


        protected virtual bool HandleUnexpectedCharacter(string input, ref List<Token> output, ref int index)
        {
            return false;
        }
    }


    public class Token
    {
        public readonly string type;
        public readonly string text;


        public Token(string type, string text = "")
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
