using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace FormulaNavigator.Core
{
    public enum FormulaNodeKind
    {
        Literal, Reference, Name, Function, Unary, Binary, Group, Array, MissingArgument
    }

    public sealed class FormulaNode
    {
        public FormulaNodeKind Kind { get; private set; }
        public string Text { get; private set; }
        public int Start { get; private set; }
        public int Length { get; private set; }
        public string Operator { get; private set; }
        public string Name { get; private set; }
        public IReadOnlyList<FormulaNode> Children { get; private set; }

        internal FormulaNode(FormulaNodeKind kind, string source, int start, int length,
            string op, string name, IList<FormulaNode> children)
        {
            Kind = kind;
            Start = start;
            Length = length;
            Text = source.Substring(start, length);
            Operator = op;
            Name = name;
            Children = new ReadOnlyCollection<FormulaNode>(children ?? new List<FormulaNode>());
        }
    }

    public sealed class FormulaParseResult
    {
        public FormulaNode Root { get; private set; }
        public IReadOnlyList<string> Diagnostics { get; private set; }
        public bool Success { get { return Diagnostics.Count == 0 && Root != null; } }

        internal FormulaParseResult(FormulaNode root, IList<string> diagnostics)
        {
            Root = root;
            Diagnostics = new ReadOnlyCollection<string>(diagnostics);
        }
    }
}
