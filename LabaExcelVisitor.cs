using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace LabaExcel
{
    class LabaExcelVisitor: LabaExcelBaseVisitor<double>
    {
        Dictionary<string, double> tableIdentifier = new Dictionary<string, double>();
        public override double VisitCompileUnit(LabaExcelParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitNumberExpr(LabaExcelParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);
            return result;
        }
        public override double VisitIdentifierExpr(LabaExcelParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            double value;
            if (tableIdentifier.TryGetValue(result.ToString(), out value))
            {
                return value;
            }
            else
            {
                return 0.0;
            }
        }
        public override double VisitParenthesizedExpr(LabaExcelParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitExponentialExpr(LabaExcelParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            Debug.WriteLine("{0} ^ {1}", left, right);
            return System.Math.Pow(left, right);
        }
        public override double VisitAdditiveExpr(LabaExcelParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabaExcelLexer.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }
        public override double VisitMultiplicativeExpr(LabaExcelParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabaExcelLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                return left * right;
            }
            else
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }
        public override double VisitEqualExpr(LabaExcelParser.EqualExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            Debug.WriteLine("{0} = {1}", left, right);
            if (left == right)
            {
                return 1;
            }
            else
                return 0;
        }
        public override double VisitLessExpr(LabaExcelParser.LessExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            Debug.WriteLine("{0} < {1}", left, right);
            if(left < right)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public override double VisitGreaterExpr(LabaExcelParser.GreaterExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            Debug.WriteLine("{0} < {1}", left, right);
            if (left > right)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
        public override double VisitMaxMinExpr(LabaExcelParser.MaxMinExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabaExcelLexer.MAX)
            {
                Debug.WriteLine("max (", left, right, ")");
                return Math.Max(left, right);
            }
            else
            {
                Debug.WriteLine("min (", left, right, ")");
                return Math.Min(left, right);
            }
        }
        public override double VisitMmaxMminExpr(LabaExcelParser.MmaxMminExprContext context)
        {
            List<double> list = new List<double>();
            int idx = 0;
            while(WalkN(context, idx) != 1.0101)
            {
                list.Add(WalkN(context, idx));
                idx++;
            }
            var sortedList = from l in list orderby l select l;
            if(context.operatorToken.Type == LabaExcelLexer.MMAX)
            {
                Debug.WriteLine("mmax(");
                for(idx = 0; idx < list.Count(); ++idx)
                {
                    Debug.WriteLine("{0}", list.ElementAt(idx));
                    if (idx != list.Count() - 1) Debug.WriteLine(" , ");
                }
                Debug.WriteLine(")");
                return sortedList.ElementAt(sortedList.Count() - 1);
            }
            else
            {
                Debug.WriteLine("mmin(");
                for(idx = 0; idx < list.Count(); ++idx)
                {
                    Debug.WriteLine("{0}", list.ElementAt(idx));
                    if (idx != list.Count() - 1) Debug.WriteLine(" , ");
                }
                Debug.WriteLine(")");
                return sortedList.ElementAt(0);
            }
        }
        private double WalkLeft(LabaExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabaExcelParser.ExpressionContext>(0));
        }
        private double WalkRight(LabaExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabaExcelParser.ExpressionContext>(1));
        }
        private double WalkN(LabaExcelParser.ExpressionContext context, int idx)
        {
            try
            {
                return Visit(context.GetRuleContext<LabaExcelParser.ExpressionContext>(idx));
            }
            catch (NullReferenceException) { return 1.0101; }
        }
    }
}
