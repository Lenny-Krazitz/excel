using System;
using System.Collections.Generic;

namespace FormulaNavigator.Core
{
    public sealed class FormulaReadBlock
    {
        public FormulaReadBlock(int row, int column, int height, int width)
        {
            Row = row;
            Column = column;
            Height = height;
            Width = width;
        }

        public int Row { get; private set; }
        public int Column { get; private set; }
        public int Height { get; private set; }
        public int Width { get; private set; }
    }

    public static class FormulaReadBlocks
    {
        public static IEnumerable<FormulaReadBlock> Enumerate(int rows, int columns, int maxCells = 4096, int maxColumns = 128)
        {
            if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows));
            if (columns < 0) throw new ArgumentOutOfRangeException(nameof(columns));
            if (maxCells < 1) throw new ArgumentOutOfRangeException(nameof(maxCells));
            if (maxColumns < 1) throw new ArgumentOutOfRangeException(nameof(maxColumns));
            if (rows == 0 || columns == 0) yield break;

            int stripWidth = Math.Min(columns, Math.Min(maxColumns, maxCells));
            int rowBandHeight = Math.Max(1, maxCells / stripWidth);
            int row = 1;
            while (true)
            {
                int height = Math.Min(rowBandHeight, rows - row + 1);
                int column = 1;
                while (true)
                {
                    int width = Math.Min(stripWidth, columns - column + 1);
                    yield return new FormulaReadBlock(row, column, height, width);
                    if (width == columns - column + 1) break;
                    column += width;
                }
                if (height == rows - row + 1) yield break;
                row += height;
            }
        }
    }
}
