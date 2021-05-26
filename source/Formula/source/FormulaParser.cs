using System.Collections.Generic;


namespace STK.Formula
{
    public class FormulaParser
    {
        public Formula Parse(List<FormulaToken> input)
        {
            int i = 0;
            return new Formula(ParseExpression(input, ref i));
        }


        private IEvaluable ParseExpression(List<FormulaToken> input, ref int i)
        {
            if (i >= input.Count)
                return null;


            IEvaluable lhs = ParseTerm(input, ref i);
            if (lhs != null && i < input.Count)
            {
                IEvaluable rhs;
                switch (input[i].type)
                {
                    case "ADDITION":
                        ++i;

                        rhs = ParseExpression(input, ref i);
                        if (rhs == null)
                        {
                            return null;
                        }

                        return new AdditionNode(lhs, rhs);


                    case "SUBTRACTION":
                        ++i;

                        rhs = ParseExpression(input, ref i);
                        if (rhs == null)
                        {
                            return null;
                        }

                        return new SubstractionNode(lhs, rhs);


                    default:
                        return HandleUnexpectedExpression(input, ref i, lhs);
                }
            }


            return lhs;
        }


        private IEvaluable ParseTerm(List<FormulaToken> input, ref int i)
        {
            if (i >= input.Count)
                return null;


            IEvaluable lhs = ParseFactor(input, ref i);
            if (lhs != null && i < input.Count)
            {
                IEvaluable rhs;
                switch (input[i].type)
                {
                    case "MULTIPLICATION":
                        ++i;

                        rhs = ParseTerm(input, ref i);
                        if (rhs == null)
                        {
                            return null;
                        }

                        return new MultiplicationNode(lhs, rhs);


                    case "DIVISION":
                        ++i;

                        rhs = ParseTerm(input, ref i);
                        if (rhs == null)
                        {
                            return null;
                        }

                        return new DivisionNode(lhs, rhs);


                    default:
                        return HandleUnexpectedTerm(input, ref i, lhs);
                }
            }


            return lhs;
        }


        private IEvaluable ParseFactor(List<FormulaToken> input, ref int i)
        {
            if (i >= input.Count)
                return null;


            FormulaToken token = input[i];


            switch (token.type)
            {
                case "CONSTANT":
                    ++i;
                    return new ConstantNode(float.Parse(token.text));


                case "VARIABLE":
                    ++i;
                    return new VariableNode(token.text);


                case "LEFT_PARENTHESIS":
                    ++i;
                    IEvaluable result = ParseExpression(input, ref i);
                    if (input[i++].type == "RIGHT_PARENTHESIS")
                    {
                        return result;
                    }

                    return null;


                default:
                    return HandleUnexpectedFactor(input, ref i);
            }
        }


        protected virtual IEvaluable HandleUnexpectedExpression(List<FormulaToken> input, ref int index, IEvaluable lhs)
        {
            return lhs;
        }


        protected virtual IEvaluable HandleUnexpectedTerm(List<FormulaToken> input, ref int index, IEvaluable lhs)
        {
            return lhs;
        }


        protected virtual IEvaluable HandleUnexpectedFactor(List<FormulaToken> input, ref int index)
        {
            return null;
        }
    }
}
