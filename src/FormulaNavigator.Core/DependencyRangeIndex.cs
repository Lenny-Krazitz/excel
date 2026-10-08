using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace FormulaNavigator.Core
{
    public sealed class IndexedDependency
    {
        public IndexedDependency(ReferenceArea area, int formulaId)
        {
            if (area == null) throw new ArgumentNullException(nameof(area));
            Area = area;
            FormulaId = formulaId;
        }

        public ReferenceArea Area { get; private set; }
        public int FormulaId { get; private set; }
    }

    // An immutable, per-sheet bounding-box tree. It indexes ranges directly, so
    // full-row and full-column references never expand into individual cells.
    public sealed class DependencyRangeIndex
    {
        private readonly Dictionary<SheetKey, Node> _roots;

        public DependencyRangeIndex(IEnumerable<IndexedDependency> dependencies)
        {
            if (dependencies == null) throw new ArgumentNullException(nameof(dependencies));
            var groups = new Dictionary<SheetKey, List<IndexedDependency>>();
            foreach (IndexedDependency dependency in dependencies)
            {
                if (dependency == null) throw new ArgumentException("Dependencies cannot contain null.", nameof(dependencies));
                var key = new SheetKey(dependency.Area.Workbook, dependency.Area.Sheet);
                List<IndexedDependency> values;
                if (!groups.TryGetValue(key, out values))
                {
                    values = new List<IndexedDependency>();
                    groups.Add(key, values);
                }
                values.Add(dependency);
            }
            _roots = new Dictionary<SheetKey, Node>();
            foreach (KeyValuePair<SheetKey, List<IndexedDependency>> group in groups)
                _roots.Add(group.Key, Node.Create(group.Value));
        }

        public IReadOnlyList<int> Find(ReferenceArea source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            Node root;
            if (!_roots.TryGetValue(new SheetKey(source.Workbook, source.Sheet), out root))
                return new ReadOnlyCollection<int>(new List<int>());
            var ids = new HashSet<int>();
            root.Find(source, ids);
            var result = new List<int>(ids);
            result.Sort();
            return new ReadOnlyCollection<int>(result);
        }

        private sealed class Node
        {
            private const int LeafSize = 16;
            private readonly int _startRow;
            private readonly int _endRow;
            private readonly int _startColumn;
            private readonly int _endColumn;
            private readonly Node _left;
            private readonly Node _right;
            private readonly IReadOnlyList<IndexedDependency> _items;

            private Node(List<IndexedDependency> items, Node left, Node right)
            {
                _left = left;
                _right = right;
                _items = items == null ? null : new ReadOnlyCollection<IndexedDependency>(items);
                bool first = true;
                int startRow = 0, endRow = 0, startColumn = 0, endColumn = 0;
                if (items != null) SetBounds(items, ref first, ref startRow, ref endRow, ref startColumn, ref endColumn);
                if (left != null) SetBounds(left, ref first, ref startRow, ref endRow, ref startColumn, ref endColumn);
                if (right != null) SetBounds(right, ref first, ref startRow, ref endRow, ref startColumn, ref endColumn);
                _startRow = startRow;
                _endRow = endRow;
                _startColumn = startColumn;
                _endColumn = endColumn;
            }

            internal static Node Create(List<IndexedDependency> values)
            {
                if (values.Count <= LeafSize) return new Node(new List<IndexedDependency>(values), null, null);
                int minRow, maxRow, minColumn, maxColumn;
                Bounds(values, out minRow, out maxRow, out minColumn, out maxColumn);
                bool splitRows = maxRow - minRow >= maxColumn - minColumn;
                values.Sort(delegate(IndexedDependency left, IndexedDependency right)
                {
                    long leftCenter = splitRows
                        ? (long)left.Area.StartRow + left.Area.EndRow : (long)left.Area.StartColumn + left.Area.EndColumn;
                    long rightCenter = splitRows
                        ? (long)right.Area.StartRow + right.Area.EndRow : (long)right.Area.StartColumn + right.Area.EndColumn;
                    return leftCenter.CompareTo(rightCenter);
                });
                int middle = values.Count / 2;
                var leftValues = values.GetRange(0, middle);
                var rightValues = values.GetRange(middle, values.Count - middle);
                return new Node(null, Create(leftValues), Create(rightValues));
            }

            internal void Find(ReferenceArea source, HashSet<int> ids)
            {
                if (!Intersects(source, _startRow, _endRow, _startColumn, _endColumn)) return;
                if (_items != null)
                {
                    for (int i = 0; i < _items.Count; i++)
                        if (Intersects(_items[i].Area, source.StartRow, source.EndRow,
                            source.StartColumn, source.EndColumn)) ids.Add(_items[i].FormulaId);
                    return;
                }
                _left.Find(source, ids);
                _right.Find(source, ids);
            }

            private static void Bounds(List<IndexedDependency> values, out int minRow, out int maxRow,
                out int minColumn, out int maxColumn)
            {
                minRow = Int32.MaxValue; maxRow = Int32.MinValue;
                minColumn = Int32.MaxValue; maxColumn = Int32.MinValue;
                for (int i = 0; i < values.Count; i++)
                {
                    ReferenceArea area = values[i].Area;
                    minRow = Math.Min(minRow, area.StartRow); maxRow = Math.Max(maxRow, area.EndRow);
                    minColumn = Math.Min(minColumn, area.StartColumn); maxColumn = Math.Max(maxColumn, area.EndColumn);
                }
            }

            private static void SetBounds(List<IndexedDependency> items, ref bool first, ref int startRow,
                ref int endRow, ref int startColumn, ref int endColumn)
            {
                for (int i = 0; i < items.Count; i++) SetBounds(items[i].Area, ref first, ref startRow, ref endRow, ref startColumn, ref endColumn);
            }

            private static void SetBounds(Node node, ref bool first, ref int startRow, ref int endRow,
                ref int startColumn, ref int endColumn)
            {
                SetBounds(node._startRow, node._endRow, node._startColumn, node._endColumn,
                    ref first, ref startRow, ref endRow, ref startColumn, ref endColumn);
            }

            private static void SetBounds(ReferenceArea area, ref bool first, ref int startRow, ref int endRow,
                ref int startColumn, ref int endColumn)
            {
                SetBounds(area.StartRow, area.EndRow, area.StartColumn, area.EndColumn,
                    ref first, ref startRow, ref endRow, ref startColumn, ref endColumn);
            }

            private static void SetBounds(int candidateStartRow, int candidateEndRow, int candidateStartColumn,
                int candidateEndColumn, ref bool first, ref int startRow, ref int endRow, ref int startColumn, ref int endColumn)
            {
                if (first)
                {
                    startRow = candidateStartRow; endRow = candidateEndRow;
                    startColumn = candidateStartColumn; endColumn = candidateEndColumn;
                    first = false;
                    return;
                }
                startRow = Math.Min(startRow, candidateStartRow); endRow = Math.Max(endRow, candidateEndRow);
                startColumn = Math.Min(startColumn, candidateStartColumn); endColumn = Math.Max(endColumn, candidateEndColumn);
            }

            private static bool Intersects(ReferenceArea area, int startRow, int endRow, int startColumn, int endColumn)
            {
                return area.StartRow <= endRow && area.EndRow >= startRow &&
                    area.StartColumn <= endColumn && area.EndColumn >= startColumn;
            }
        }

        private sealed class SheetKey : IEquatable<SheetKey>
        {
            private readonly string _workbook;
            private readonly string _sheet;

            internal SheetKey(string workbook, string sheet)
            {
                _workbook = workbook ?? String.Empty;
                _sheet = sheet ?? String.Empty;
            }

            public bool Equals(SheetKey other)
            {
                return other != null && String.Equals(_workbook, other._workbook, StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(_sheet, other._sheet, StringComparison.OrdinalIgnoreCase);
            }

            public override bool Equals(object obj) { return Equals(obj as SheetKey); }
            public override int GetHashCode()
            {
                return StringComparer.OrdinalIgnoreCase.GetHashCode(_workbook) * 397 ^
                    StringComparer.OrdinalIgnoreCase.GetHashCode(_sheet);
            }
        }
    }
}
