using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace FormulaNavigator.Core
{
    public enum FormulaNavigationNodeKind
    {
        Operation, Reference, CalculatedReference
    }

    // A view of the parser's AST, not a replacement for it: evaluation and highlighting
    // keep the complete original expression, including constants and omitted arguments.
    public sealed class FormulaNavigationNode
    {
        public FormulaNode Expression { get; private set; }
        public FormulaNavigationNodeKind Kind { get; private set; }
        public int ArgumentIndex { get; private set; }
        public IReadOnlyList<FormulaNavigationNode> Children { get; private set; }

        private FormulaNavigationNode(FormulaNode expression, FormulaNavigationNodeKind kind,
            int argumentIndex, IList<FormulaNavigationNode> children)
        {
            Expression = expression;
            Kind = kind;
            ArgumentIndex = argumentIndex;
            Children = new ReadOnlyCollection<FormulaNavigationNode>(children ?? new List<FormulaNavigationNode>());
        }

        public static FormulaNavigationNode Create(FormulaNode expression)
        {
            return Project(expression, -1);
        }

        private static FormulaNavigationNode Project(FormulaNode expression, int argumentIndex)
        {
            if (expression == null || expression.Kind == FormulaNodeKind.Literal ||
                expression.Kind == FormulaNodeKind.MissingArgument) return null;

            // A range/spill is one reference, rather than unrelated endpoint/anchor leaves.
            if (expression.Kind == FormulaNodeKind.Reference || expression.Kind == FormulaNodeKind.Name ||
                expression.Kind == FormulaNodeKind.Binary && expression.Operator == ":" ||
                expression.Kind == FormulaNodeKind.Unary && expression.Operator == "#")
                return new FormulaNavigationNode(expression, FormulaNavigationNodeKind.Reference, argumentIndex, null);

            var children = new List<FormulaNavigationNode>();
            for (int i = 0; i < expression.Children.Count; i++)
            {
                var child = Project(expression.Children[i], i);
                if (child != null) children.Add(child);
            }

            if (IsCalculatedReference(expression))
            {
                // INDIRECT("A"&2) still leads to a cell even with no reference tokens in
                // its arguments. Keep a dedicated leaf for the address resolved by Excel.
                children.Add(new FormulaNavigationNode(expression,
                    FormulaNavigationNodeKind.CalculatedReference, -1, null));
            }

            // SUM(1,2), arithmetic on literals, strings and literal arrays do not become
            // terminal operation nodes. Their source text stays in the nearest retained node.
            if (children.Count == 0) return null;
            return new FormulaNavigationNode(expression, FormulaNavigationNodeKind.Operation, argumentIndex, children);
        }

        private static bool IsCalculatedReference(FormulaNode expression)
        {
            return expression.Kind == FormulaNodeKind.Function &&
                (String.Equals(expression.Name, "OFFSET", StringComparison.OrdinalIgnoreCase) ||
                 String.Equals(expression.Name, "INDIRECT", StringComparison.OrdinalIgnoreCase));
        }
    }
}
