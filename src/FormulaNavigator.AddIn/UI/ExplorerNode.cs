using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using FormulaNavigator.AddIn.Excel;
using FormulaNavigator.Core;

namespace FormulaNavigator.AddIn.UI
{
    public sealed class ExplorerNode : INotifyPropertyChanged
    {
        private bool inspected;
        private readonly bool calculatedReference;
        private string value = "…";
        private string address = "";
        private string note = "";
        public InspectionContext Context { get; private set; }
        public FormulaNode Expression { get; private set; }
        public string Label { get; private set; }
        public string Argument { get; private set; }
        public string Value { get { return value; } }
        public string Address { get { return address; } }
        public string Note { get { return note; } }
        public ExcelLocation Target { get; private set; }
        public ObservableCollection<ExplorerNode> Children { get; private set; }
        public event PropertyChangedEventHandler PropertyChanged;

        private ExplorerNode(InspectionContext context, FormulaNavigationNode projection, string argument)
        {
            Context = context;
            Expression = projection == null ? null : projection.Expression;
            Argument = argument ?? "";
            Children = new ObservableCollection<ExplorerNode>();
            if (projection == null)
            {
                Label = context.Location.Address;
                value = context.DisplayValue;
                Target = context.Location;
                address = Target.ToString();
                inspected = true;
                return;
            }
            calculatedReference = projection.Kind == FormulaNavigationNodeKind.CalculatedReference;
            Label = calculatedReference ? "Вычисляемая ссылка" : MakeLabel(Expression);
            foreach (var child in projection.Children)
                Children.Add(new ExplorerNode(context, child,
                    child.Kind == FormulaNavigationNodeKind.CalculatedReference
                        ? "результат " + Expression.Name : ArgumentName(Expression, child.ArgumentIndex)));
        }

        public static ExplorerNode CreateRoot(InspectionContext context, FormulaNode expression)
        {
            var root = new ExplorerNode(context, null, "исходная ячейка");
            var projection = FormulaNavigationNode.Create(expression);
            if (projection != null) root.Children.Add(new ExplorerNode(context, projection, "формула"));
            return root;
        }

        public void Inspect(ExcelGateway gateway)
        {
            if (inspected) return;
            // Set this only after a successful call: transient COM errors can be retried.
            var result = gateway.Inspect(Context, Expression);
            value = result.DisplayValue ?? "";
            note = result.Note ?? "";
            Target = result.Target;
            address = Target == null ? "" : Target.ToString();
            if (calculatedReference)
            {
                Label = Target == null ? "Ссылка не определена" : Target.Sheet + "!" + Target.Address;
                if (Target == null && string.IsNullOrEmpty(note)) note = "Вычислить адрес этой ссылки не удалось.";
                Notify("Label");
            }
            inspected = true;
            Notify("Value"); Notify("Address"); Notify("Note");
        }

        private void Notify(string property)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(property));
        }

        private static string MakeLabel(FormulaNode node)
        {
            if (node.Kind == FormulaNodeKind.Function) return node.Name;
            if (node.Kind == FormulaNodeKind.MissingArgument) return "(пропущен)";
            var text = node.Text.Replace("\r", " ").Replace("\n", " ");
            return text.Length > 110 ? text.Substring(0, 107) + "…" : text;
        }

        private static string ArgumentName(FormulaNode node, int index)
        {
            if (node.Kind == FormulaNodeKind.Binary) return index == 0 ? "слева" : "справа";
            if (node.Kind != FormulaNodeKind.Function) return "";
            string[] names;
            if (Arguments.TryGetValue(node.Name ?? "", out names) && index < names.Length)
                return names[index];
            return "аргумент " + (index + 1);
        }

        private static readonly Dictionary<string, string[]> Arguments = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "IF", new[] { "logical_test", "value_if_true", "value_if_false" } },
            { "IFERROR", new[] { "value", "value_if_error" } },
            { "IFNA", new[] { "value", "value_if_na" } },
            { "OFFSET", new[] { "reference", "rows", "cols", "height", "width" } },
            { "INDIRECT", new[] { "ref_text", "a1" } },
            { "INDEX", new[] { "reference", "row_num", "column_num", "area_num" } },
            { "MATCH", new[] { "lookup_value", "lookup_array", "match_type" } },
            { "XLOOKUP", new[] { "lookup_value", "lookup_array", "return_array", "if_not_found", "match_mode", "search_mode" } },
            { "VLOOKUP", new[] { "lookup_value", "table_array", "col_index_num", "range_lookup" } },
            { "SUMIF", new[] { "range", "criteria", "sum_range" } },
            { "COUNTIF", new[] { "range", "criteria" } }
        };
    }
}
