using System;
using System.Globalization;

namespace FormulaNavigator.Core
{
    public sealed class ReferenceArea
    {
        public ReferenceArea(string workbook, string sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            Workbook = workbook;
            Sheet = sheet;
            StartRow = startRow;
            EndRow = endRow;
            StartColumn = startColumn;
            EndColumn = endColumn;
        }

        public string Workbook { get; private set; }
        public string Sheet { get; private set; }
        public int StartRow { get; private set; }
        public int EndRow { get; private set; }
        public int StartColumn { get; private set; }
        public int EndColumn { get; private set; }

        public bool Intersects(ReferenceArea other)
        {
            if (other == null || !String.Equals(Workbook, other.Workbook, StringComparison.OrdinalIgnoreCase)
                || !String.Equals(Sheet, other.Sheet, StringComparison.OrdinalIgnoreCase)) return false;
            return StartRow <= other.EndRow && EndRow >= other.StartRow
                && StartColumn <= other.EndColumn && EndColumn >= other.StartColumn;
        }
    }

    public static class ReferenceParser
    {
        public const int MaxRows = 1048576;
        public const int MaxColumns = 16384;

        public static bool TryParse(string text, string defaultWorkbook, string defaultSheet, out ReferenceArea area)
        {
            area = null;
            if (String.IsNullOrWhiteSpace(text)) return false;
            text = text.Trim();
            string workbook = defaultWorkbook;
            string sheet = defaultSheet;
            int bang = FindQualifierBang(text);
            if (bang >= 0)
            {
                if (!TryParseQualifier(text.Substring(0, bang), out workbook, out sheet)) return false;
                workbook = workbook ?? defaultWorkbook;
                text = text.Substring(bang + 1);
            }
            int colon = text.IndexOf(':');
            string first = colon < 0 ? text : text.Substring(0, colon);
            string second = colon < 0 ? first : text.Substring(colon + 1);
            if (colon >= 0 && (second.Length == 0 || second.IndexOf(':') >= 0)) return false;
            int sr, er, sc, ec, type1, type2;
            if (!TryParsePart(first, out sr, out er, out sc, out ec, out type1)) return false;
            int sr2, er2, sc2, ec2;
            if (!TryParsePart(second, out sr2, out er2, out sc2, out ec2, out type2) || type1 != type2) return false;
            area = new ReferenceArea(workbook, sheet, Math.Min(sr, sr2), Math.Max(er, er2),
                Math.Min(sc, sc2), Math.Max(ec, ec2));
            return true;
        }

        private static int FindQualifierBang(string value)
        {
            bool quoted = false;
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] == '\'')
                {
                    if (quoted && i + 1 < value.Length && value[i + 1] == '\'') { i++; continue; }
                    quoted = !quoted;
                }
                else if (value[i] == '!' && !quoted) return i;
            }
            return -1;
        }

        private static bool TryParseQualifier(string value, out string workbook, out string sheet)
        {
            workbook = null;
            sheet = null;
            if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'')
                value = value.Substring(1, value.Length - 2).Replace("''", "'");
            if (value.Length > 2 && value[0] == '[')
            {
                int close = value.IndexOf(']');
                if (close <= 1 || close == value.Length - 1) return false;
                workbook = value.Substring(1, close - 1);
                sheet = value.Substring(close + 1);
            }
            else sheet = value;
            return !String.IsNullOrEmpty(sheet);
        }

        // type: 0 cell, 1 entire column, 2 entire row
        private static bool TryParsePart(string value, out int startRow, out int endRow,
            out int startColumn, out int endColumn, out int type)
        {
            startRow = endRow = startColumn = endColumn = type = 0;
            if (String.IsNullOrEmpty(value)) return false;
            int i = 0;
            if (value[i] == '$') i++;
            int lettersStart = i;
            while (i < value.Length && Char.IsLetter(value[i])) i++;
            int lettersLength = i - lettersStart;
            string letters = value.Substring(lettersStart, lettersLength);
            if (i < value.Length && value[i] == '$') i++;
            int digitsStart = i;
            while (i < value.Length && Char.IsDigit(value[i])) i++;
            int digitsLength = i - digitsStart;
            if (i != value.Length || (lettersLength == 0 && digitsLength == 0)) return false;
            int col = 0;
            int row = 0;
            if (lettersLength > 0 && !TryColumn(letters, out col)) return false;
            if (digitsLength > 0 && !Int32.TryParse(value.Substring(digitsStart, digitsLength), NumberStyles.None,
                CultureInfo.InvariantCulture, out row)) return false;
            if (lettersLength > 0 && digitsLength > 0)
            {
                if (row < 1 || row > MaxRows) return false;
                startRow = endRow = row; startColumn = endColumn = col; type = 0; return true;
            }
            if (lettersLength > 0)
            {
                startRow = 1; endRow = MaxRows; startColumn = endColumn = col; type = 1; return true;
            }
            if (row < 1 || row > MaxRows) return false;
            startRow = endRow = row; startColumn = 1; endColumn = MaxColumns; type = 2; return true;
        }

        private static bool TryColumn(string value, out int column)
        {
            column = 0;
            if (value.Length > 3) return false;
            foreach (char c in value.ToUpperInvariant())
            {
                if (c < 'A' || c > 'Z') return false;
                column = checked(column * 26 + (c - 'A' + 1));
            }
            return column >= 1 && column <= MaxColumns;
        }
    }
}
