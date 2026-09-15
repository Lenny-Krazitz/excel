using System;
using System.Collections.Generic;

namespace FormulaNavigator.AddIn.Excel
{
    public sealed class ExcelLocation
    {
        public ExcelLocation(string workbook, string sheet, string address)
        {
            if (String.IsNullOrWhiteSpace(workbook)) throw new ArgumentException("A workbook name is required.", "workbook");
            if (String.IsNullOrWhiteSpace(sheet)) throw new ArgumentException("A worksheet name is required.", "sheet");
            if (String.IsNullOrWhiteSpace(address)) throw new ArgumentException("An A1 address is required.", "address");
            Workbook = workbook;
            Sheet = sheet;
            Address = address;
            Key = workbook + "|" + sheet + "|" + address;
        }

        public string Workbook { get; private set; }
        public string Sheet { get; private set; }
        public string Address { get; private set; }
        public string Key { get; private set; }

        public override string ToString()
        {
            return "[" + Workbook + "]" + Sheet + "!" + Address;
        }
    }

    public sealed class InspectionContext
    {
        public InspectionContext(ExcelLocation location, string formula, string displayValue, bool hasFormula)
        {
            if (location == null) throw new ArgumentNullException("location");
            Location = location;
            Formula = formula ?? String.Empty;
            DisplayValue = displayValue ?? String.Empty;
            HasFormula = hasFormula;
        }

        public ExcelLocation Location { get; private set; }
        public string Formula { get; private set; }
        public string DisplayValue { get; private set; }
        public bool HasFormula { get; private set; }
    }

    public sealed class NodeInspection
    {
        public NodeInspection(string displayValue, string note, ExcelLocation target)
        {
            DisplayValue = displayValue ?? String.Empty;
            Note = note ?? String.Empty;
            Target = target;
        }

        public string DisplayValue { get; private set; }
        public string Note { get; private set; }
        public ExcelLocation Target { get; private set; }
    }

    public sealed class DependentCell
    {
        public DependentCell(InspectionContext context, string note)
        {
            if (context == null) throw new ArgumentNullException("context");
            Context = context;
            Note = note ?? String.Empty;
        }

        public InspectionContext Context { get; private set; }
        public string Note { get; private set; }
    }

    public sealed class DependencyResult
    {
        public DependencyResult(IReadOnlyList<DependentCell> cells, IReadOnlyList<string> warnings, int formulaCount,
            bool indexReused = false, long elapsedMilliseconds = 0)
        {
            Cells = cells ?? new List<DependentCell>().AsReadOnly();
            Warnings = warnings ?? new List<string>().AsReadOnly();
            FormulaCount = formulaCount;
            IndexReused = indexReused;
            ElapsedMilliseconds = elapsedMilliseconds;
        }

        public IReadOnlyList<DependentCell> Cells { get; private set; }
        public IReadOnlyList<string> Warnings { get; private set; }
        public int FormulaCount { get; private set; }
        public bool IndexReused { get; private set; }
        public long ElapsedMilliseconds { get; private set; }
    }
}
