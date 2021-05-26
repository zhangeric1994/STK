using System.Collections.Generic;


namespace STK.Formula
{
    public interface IEvaluable
    {
        float Evaluate(params IDictionary<string, float>[] variableDictionaries);
    }


    public readonly struct Formula : IEvaluable
    {
        public static Formula NONE = new Formula(null);


        private readonly IEvaluable startNode;


        public static Formula operator +(Formula lhs, Formula rhs)
        {
            if (lhs.startNode == null)
            {
                return rhs;
            }


            if (rhs.startNode == null)
            {
                return lhs;
            }


            return new Formula(new AdditionNode(lhs.startNode, rhs.startNode));
        }

        public static Formula operator -(Formula lhs, Formula rhs)
        {
            if (lhs.startNode == null)
            {
                return new Formula(new NegationNode(rhs.startNode));
            }


            if (rhs.startNode == null)
            {
                return lhs ;
            }


            return new Formula(new SubstractionNode(lhs.startNode, rhs.startNode));
        }

        public static Formula operator *(Formula lhs, Formula rhs)
        {
            if (lhs.startNode == null)
            {
                return NONE;
            }


            if (rhs.startNode == null)
            {
                return NONE;
            }


            return new Formula(new MultiplicationNode(lhs.startNode, rhs.startNode));
        }

        public static Formula operator /(Formula lhs, Formula rhs)
        {
            if (lhs.startNode == null)
            {
                return NONE;
            }


            if (rhs.startNode == null)
            {
                return NONE;
            }


            return new Formula(new DivisionNode(lhs.startNode, rhs.startNode));
        }


        public Formula(IEvaluable startNode)
        {
            this.startNode = startNode;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            return startNode == null ? 0 : startNode.Evaluate(variableDictionaries);
        }
    }


    public readonly struct ConstantNode : IEvaluable
    {
        private readonly float value;


        public ConstantNode(float value)
        {
            this.value = value;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            return value;
        }
    }


    public readonly struct VariableNode : IEvaluable
    {
        private readonly string name;


        public VariableNode(string name)
        {
            this.name = name;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            float result = 0;
            foreach (IDictionary<string, float> variableSet in variableDictionaries)
            {
                if (variableSet != null)
                {
                    if (variableSet.TryGetValue(name, out float value))
                    {
                        result += value;
                    }
                }
            }

            return result;
        }
    }


    public readonly struct NegationNode : IEvaluable
    {
        private readonly IEvaluable node;


        public NegationNode(IEvaluable node)
        {
            this.node = node;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            return -node.Evaluate(variableDictionaries);
        }
    }


    public readonly struct AdditionNode : IEvaluable
    {
        private readonly IEvaluable lhs;
        private readonly IEvaluable rhs;


        public AdditionNode(IEvaluable lhs, IEvaluable rhs)
        {
            this.lhs = lhs;
            this.rhs = rhs;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            return lhs.Evaluate(variableDictionaries) + rhs.Evaluate(variableDictionaries);
        }
    }


    public readonly struct SubstractionNode : IEvaluable
    {
        private readonly IEvaluable minuend;
        private readonly IEvaluable subtrahend;


        public SubstractionNode(IEvaluable minuend, IEvaluable subtrahend)
        {
            this.minuend = minuend;
            this.subtrahend = subtrahend;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            return minuend.Evaluate(variableDictionaries) - subtrahend.Evaluate(variableDictionaries);
        }
    }


    public readonly struct MultiplicationNode : IEvaluable
    {
        private readonly IEvaluable lhs;
        private readonly IEvaluable rhs;


        public MultiplicationNode(IEvaluable lhs, IEvaluable rhs)
        {
            this.lhs = lhs;
            this.rhs = rhs;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            return lhs.Evaluate(variableDictionaries) * rhs.Evaluate(variableDictionaries);
        }
    }


    public readonly struct DivisionNode : IEvaluable
    {
        private readonly IEvaluable dividend;
        private readonly IEvaluable divisor;


        public DivisionNode(IEvaluable dividend, IEvaluable divisor)
        {
            this.dividend = dividend;
            this.divisor = divisor;
        }


        public float Evaluate(params IDictionary<string, float>[] variableDictionaries)
        {
            return dividend.Evaluate(variableDictionaries) / divisor.Evaluate(variableDictionaries);
        }
    }
}
