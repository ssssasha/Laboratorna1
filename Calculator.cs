using System;
using Antlr4.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LabaExcel
{
    public static class Calculator
    {
        public static double Evaluate(string expression)
        {
            var lexer = new LabaExcelLexer(new AntlrInputStream(expression));
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());
            var tokens = new CommonTokenStream(lexer);
            var parser = new LabaExcelParser(tokens);
            var tree = parser.compileUnit();
            var visitor = new LabaExcelVisitor ();
            return visitor.Visit(tree);
        }
    }
}
