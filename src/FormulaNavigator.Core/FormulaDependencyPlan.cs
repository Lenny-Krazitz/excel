using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;

namespace FormulaNavigator.Core
{
    // A conservative projection of a formula AST. It only reports references whose
    // addresses can be determined without consulting Excel.
    public sealed class FormulaDependencyPlan
    {
        private FormulaDependencyPlan(IList<ReferenceArea> references, IList<FormulaNode> dynamicNodes,
            IList<string> warnings)
        {
            References = new ReadOnlyCollection<ReferenceArea>(references);
            DynamicNodes = new ReadOnlyCollection<FormulaNode>(dynamicNodes);
            Warnings = new ReadOnlyCollection<string>(warnings);
        }

        public IReadOnlyList<ReferenceArea> References { get; private set; }
        public IReadOnlyList<FormulaNode> DynamicNodes { get; private set; }
        public IReadOnlyList<string> Warnings { get; private set; }

        public static FormulaDependencyPlan Create(FormulaNode root, string workbook, string sheet)
        {
            var builder = new Builder(workbook, sheet);
            builder.Visit(root);
            return new FormulaDependencyPlan(builder.References, builder.DynamicNodes, builder.Warnings);
        }

        private sealed class Builder
        {
            // Intersections of two unions can otherwise create a Cartesian product of
            // temporary areas. These limits are shared by the whole formula plan.
            private const int MaxStaticResolutionWork = 2048;
            private const int MaxStaticResolutionAreas = 4096;
            private readonly string _workbook;
            private readonly string _sheet;
            private int _staticResolutionWork;
            private int _staticResolutionAreas;
            private bool _staticResolutionCapped;

            internal Builder(string workbook, string sheet)
            {
                _workbook = workbook;
                _sheet = sheet;
                References = new List<ReferenceArea>();
                DynamicNodes = new List<FormulaNode>();
                Warnings = new List<string>();
            }

            internal List<ReferenceArea> References { get; private set; }
            internal List<FormulaNode> DynamicNodes { get; private set; }
            internal List<string> Warnings { get; private set; }

            internal void Visit(FormulaNode node)
            {
                if (node == null || node.Kind == FormulaNodeKind.Literal || node.Kind == FormulaNodeKind.MissingArgument)
                    return;

                if (node.Kind == FormulaNodeKind.Reference)
                {
                    AddReference(node.Text);
                    return;
                }

                if (node.Kind == FormulaNodeKind.Name || IsSpill(node))
                {
                    AddDynamic(node);
                    return;
                }

                if (IsOffset(node))
                {
                    ReferenceArea offsetArea;
                    if (TryResolveOffset(node, out offsetArea)) AddArea(offsetArea, node.Text);
                    else AddDynamic(node);
                    VisitChildren(node);
                    return;
                }

                if (IsIndirect(node))
                {
                    AddDynamic(node);
                    VisitChildren(node);
                    return;
                }

                if (node.Kind == FormulaNodeKind.Group)
                {
                    Visit(node.Children.Count == 0 ? null : node.Children[0]);
                    return;
                }

                if (node.Kind == FormulaNodeKind.Binary && node.Operator == ":")
                {
                    VisitRange(node);
                    return;
                }

                if (node.Kind == FormulaNodeKind.Binary && node.Operator == " ")
                {
                    VisitIntersection(node);
                    return;
                }

                // A union has independently useful operands. Traverse them even if a
                // name or another dynamic expression is mixed into the union.
                VisitChildren(node);
            }

            private void VisitRange(FormulaNode node)
            {
                ReferenceArea area;
                // Parsing the whole source span deliberately propagates a qualifier on
                // the left endpoint (Sheet1!A1:B2) to the right endpoint.
                if (ReferenceParser.TryParse(node.Text, _workbook, _sheet, out area))
                {
                    AddArea(area, node.Text);
                    return;
                }
                AddDynamic(node);
            }

            private void VisitIntersection(FormulaNode node)
            {
                if (node.Children.Count != 2)
                {
                    AddDynamic(node);
                    return;
                }
                List<ReferenceArea> left;
                List<ReferenceArea> right;
                if (!TryResolveStatic(node.Children[0], out left) || !TryResolveStatic(node.Children[1], out right))
                {
                    if (_staticResolutionCapped) return;
                    AddDynamic(node);
                    return;
                }
                if (!TryReserveResolutionWork((long)left.Count * right.Count)) return;
                for (int i = 0; i < left.Count; i++)
                {
                    for (int j = 0; j < right.Count; j++)
                    {
                        ReferenceArea intersection = Intersect(left[i], right[j]);
                        if (intersection != null)
                        {
                            if (!TryReserveResolutionArea()) return;
                            AddArea(intersection, node.Text);
                        }
                    }
                }
                VisitStaticOffsetArguments(node);
            }

            private bool TryResolveStatic(FormulaNode node, out List<ReferenceArea> areas)
            {
                areas = new List<ReferenceArea>();
                if (node == null) return false;
                if (IsOffset(node))
                {
                    ReferenceArea offsetArea;
                    if (!TryResolveOffset(node, out offsetArea)) return false;
                    return TryAddResolvedArea(areas, offsetArea);
                }
                if (node.Kind == FormulaNodeKind.Reference)
                {
                    ReferenceArea area;
                    if (!ReferenceParser.TryParse(node.Text, _workbook, _sheet, out area)) return false;
                    if (!TryAddResolvedArea(areas, area)) return false;
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Group && node.Children.Count == 1)
                    return TryResolveStatic(node.Children[0], out areas);
                if (node.Kind != FormulaNodeKind.Binary) return false;
                if (node.Operator == ":")
                {
                    ReferenceArea area;
                    if (!ReferenceParser.TryParse(node.Text, _workbook, _sheet, out area)) return false;
                    if (!TryAddResolvedArea(areas, area)) return false;
                    return true;
                }
                if (node.Operator == "," || node.Operator == ";")
                {
                    List<ReferenceArea> left;
                    List<ReferenceArea> right;
                    if (!TryResolveStatic(node.Children[0], out left) || !TryResolveStatic(node.Children[1], out right)) return false;
                    for (int i = 0; i < left.Count; i++) if (!TryAddResolvedArea(areas, left[i])) return false;
                    for (int i = 0; i < right.Count; i++) if (!TryAddResolvedArea(areas, right[i])) return false;
                    return true;
                }
                if (node.Operator == " ")
                {
                    List<ReferenceArea> left;
                    List<ReferenceArea> right;
                    if (!TryResolveStatic(node.Children[0], out left) || !TryResolveStatic(node.Children[1], out right)) return false;
                    if (!TryReserveResolutionWork((long)left.Count * right.Count)) return false;
                    for (int i = 0; i < left.Count; i++)
                    {
                        for (int j = 0; j < right.Count; j++)
                        {
                            ReferenceArea intersection = Intersect(left[i], right[j]);
                            if (intersection != null && !TryAddResolvedArea(areas, intersection)) return false;
                        }
                    }
                    return true;
                }
                return false;
            }

            private void VisitChildren(FormulaNode node)
            {
                for (int i = 0; i < node.Children.Count; i++) Visit(node.Children[i]);
            }

            // A static intersection normally suppresses its endpoint dependencies.
            // OFFSET is different: its base is a real precedent even when its shifted
            // result can also be indexed statically.
            private void VisitStaticOffsetArguments(FormulaNode node)
            {
                if (node == null) return;
                if (IsOffset(node))
                {
                    ReferenceArea ignored;
                    if (TryResolveOffset(node, out ignored))
                    {
                        VisitChildren(node);
                        return;
                    }
                }
                for (int i = 0; i < node.Children.Count; i++) VisitStaticOffsetArguments(node.Children[i]);
            }

            private void AddReference(string text)
            {
                ReferenceArea area;
                if (ReferenceParser.TryParse(text, _workbook, _sheet, out area)) AddArea(area, text);
                else AddWarning("Reference could not be indexed: " + text);
            }

            private void AddArea(ReferenceArea area, string text)
            {
                if (IsExternal(area))
                {
                    AddWarning("External reference was excluded: " + text);
                    return;
                }
                References.Add(area);
            }

            private bool IsExternal(ReferenceArea area)
            {
                return !String.IsNullOrEmpty(area.Workbook) &&
                    !String.Equals(area.Workbook, _workbook, StringComparison.OrdinalIgnoreCase);
            }

            private void AddDynamic(FormulaNode node)
            {
                DynamicNodes.Add(node);
            }

            private void AddWarning(string warning)
            {
                if (!Warnings.Contains(warning)) Warnings.Add(warning);
            }

            private bool TryAddResolvedArea(List<ReferenceArea> areas, ReferenceArea area)
            {
                if (!TryReserveResolutionArea()) return false;
                areas.Add(area);
                return true;
            }

            private bool TryReserveResolutionWork(long work)
            {
                if (work > MaxStaticResolutionWork - _staticResolutionWork)
                {
                    CapStaticResolution();
                    return false;
                }
                _staticResolutionWork += (int)work;
                return true;
            }

            private bool TryReserveResolutionArea()
            {
                if (_staticResolutionAreas >= MaxStaticResolutionAreas)
                {
                    CapStaticResolution();
                    return false;
                }
                _staticResolutionAreas++;
                return true;
            }

            private void CapStaticResolution()
            {
                if (_staticResolutionCapped) return;
                _staticResolutionCapped = true;
                AddWarning("Static dependency resolution was capped; partial dependencies were skipped.");
            }

            private bool TryResolveOffset(FormulaNode node, out ReferenceArea area)
            {
                area = null;
                if (node.Children.Count < 3 || node.Children.Count > 5) return false;
                ReferenceArea basis;
                if (!TryResolveOffsetBasis(node.Children[0], out basis)) return false;
                int rowOffset;
                int columnOffset;
                if (!TryLiteralInteger(node.Children[1], out rowOffset) ||
                    !TryLiteralInteger(node.Children[2], out columnOffset)) return false;
                int height = basis.EndRow - basis.StartRow + 1;
                int width = basis.EndColumn - basis.StartColumn + 1;
                if (node.Children.Count >= 4 && (!TryLiteralPositiveInteger(node.Children[3], out height))) return false;
                if (node.Children.Count == 5 && (!TryLiteralPositiveInteger(node.Children[4], out width))) return false;

                long startRow = (long)basis.StartRow + rowOffset;
                long startColumn = (long)basis.StartColumn + columnOffset;
                long endRow = startRow + height - 1;
                long endColumn = startColumn + width - 1;
                if (startRow < 1 || startColumn < 1 || endRow > ReferenceParser.MaxRows || endColumn > ReferenceParser.MaxColumns)
                    return false;
                area = new ReferenceArea(basis.Workbook, basis.Sheet, (int)startRow, (int)endRow,
                    (int)startColumn, (int)endColumn);
                return true;
            }

            private bool TryResolveOffsetBasis(FormulaNode node, out ReferenceArea area)
            {
                area = null;
                if (node == null) return false;
                if (node.Kind == FormulaNodeKind.Reference)
                    return ReferenceParser.TryParse(node.Text, _workbook, _sheet, out area);
                return node.Kind == FormulaNodeKind.Binary && node.Operator == ":" &&
                    ReferenceParser.TryParse(node.Text, _workbook, _sheet, out area);
            }

            private static bool TryLiteralPositiveInteger(FormulaNode node, out int value)
            {
                return TryLiteralInteger(node, out value) && value > 0;
            }

            private static bool TryLiteralInteger(FormulaNode node, out int value)
            {
                value = 0;
                if (node == null) return false;
                if (node.Kind == FormulaNodeKind.Literal)
                    return Int32.TryParse(node.Text, NumberStyles.None, CultureInfo.InvariantCulture, out value);
                if (node.Kind != FormulaNodeKind.Unary || (node.Operator != "+" && node.Operator != "-") || node.Children.Count != 1)
                    return false;
                int unsigned;
                if (!TryLiteralInteger(node.Children[0], out unsigned)) return false;
                if (node.Operator == "+") { value = unsigned; return true; }
                if (unsigned == Int32.MinValue) return false;
                value = -unsigned;
                return true;
            }

            private static bool IsOffset(FormulaNode node)
            {
                return node.Kind == FormulaNodeKind.Function &&
                    String.Equals(node.Name, "OFFSET", StringComparison.OrdinalIgnoreCase);
            }

            private static bool IsIndirect(FormulaNode node)
            {
                return node.Kind == FormulaNodeKind.Function &&
                    String.Equals(node.Name, "INDIRECT", StringComparison.OrdinalIgnoreCase);
            }

            private static bool IsSpill(FormulaNode node)
            {
                return node.Kind == FormulaNodeKind.Unary && node.Operator == "#";
            }

            private static ReferenceArea Intersect(ReferenceArea left, ReferenceArea right)
            {
                if (!String.Equals(left.Workbook, right.Workbook, StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(left.Sheet, right.Sheet, StringComparison.OrdinalIgnoreCase)) return null;
                int startRow = Math.Max(left.StartRow, right.StartRow);
                int endRow = Math.Min(left.EndRow, right.EndRow);
                int startColumn = Math.Max(left.StartColumn, right.StartColumn);
                int endColumn = Math.Min(left.EndColumn, right.EndColumn);
                return startRow <= endRow && startColumn <= endColumn
                    ? new ReferenceArea(left.Workbook, left.Sheet, startRow, endRow, startColumn, endColumn) : null;
            }
        }
    }
}
