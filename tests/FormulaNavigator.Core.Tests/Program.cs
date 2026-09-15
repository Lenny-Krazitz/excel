using System;
using System.Collections.Generic;
using System.IO;
using FormulaNavigator.Core;

namespace FormulaNavigator.Core.Tests
{
    internal static class Program
    {
        private static int _failures;

        private static int Main(string[] args)
        {
            if (args.Length == 2 && args[0] == "--verify-xll") return VerifyXll(args[1]);
            if (args.Length == 2 && args[0] == "--verify-addin") return ManagedAddInVerifier.Verify(args[1]);
            if (args.Length != 0)
            {
                Console.Error.WriteLine("Usage: FormulaNavigator.Core.Tests [--verify-xll path | --verify-addin path]");
                return 2;
            }
            Run("spans and nested functions", SpansAndNestedFunctions);
            Run("quoted and cross-sheet references", QuotedReferences);
            Run("mixed absolute reference lexing", MixedAbsoluteReferenceLexing);
            Run("repeated expressions", RepeatedExpressions);
            Run("strings and scientific literals", StringAndNumericLiterals);
            Run("missing arguments and arrays", MissingArgumentsAndArrays);
            Run("structured references", StructuredReferences);
            Run("malformed formulas", MalformedFormulas);
            Run("parser nesting and token bounds", ParserBounds);
            Run("reference areas and bounds", ReferenceAreasAndBounds);
            Run("operator precedence", OperatorPrecedence);
            Run("navigation tree cell leaves", NavigationCellLeaves);
            Run("navigation tree ranges and spans", NavigationRangesAndSpans);
            Run("navigation tree calculated references", NavigationCalculatedReferences);
            Run("dependency plan static ranges and intersections", DependencyPlanStaticRangesAndIntersections);
            Run("dependency plan static OFFSET", DependencyPlanStaticOffset);
            Run("dependency plan dynamic nodes and external references", DependencyPlanDynamicAndExternalReferences);
            Run("dependency plan intersection resolution cap", DependencyPlanIntersectionResolutionCap);
            Run("dependency range index", DependencyRangeIndexTests);
            Run("formula read block coverage", FormulaReadBlockCoverage);
            Console.WriteLine(_failures == 0 ? "All FormulaNavigator.Core tests passed." : _failures + " test(s) failed.");
            return _failures == 0 ? 0 : 1;
        }

        private static int VerifyXll(string path)
        {
            string error;
            if (XllVerifier.TryVerify(path, out error))
            {
                Console.WriteLine("XLL verification passed: " + path);
                return 0;
            }
            Console.Error.WriteLine("XLL verification failed: " + error);
            return 1;
        }

        private static void SpansAndNestedFunctions()
        {
            const string formula = "=IF(OFFSET(A1,1,2)>0,SUM('O''Brien'!$A$1:$B$2),0)";
            FormulaParseResult result = FormulaParser.Parse(formula);
            Assert(result.Success, Diagnostics(result));
            AssertNode(result.Root, FormulaNodeKind.Function, formula.Substring(1), 1, "IF");
            AssertEqual(3, result.Root.Children.Count, "IF argument count");
            FormulaNode comparison = result.Root.Children[0];
            AssertEqual(FormulaNodeKind.Binary, comparison.Kind, "comparison kind");
            AssertEqual(">", comparison.Operator, "comparison operator");
            FormulaNode offset = comparison.Children[0];
            AssertNode(offset, FormulaNodeKind.Function, "OFFSET(A1,1,2)", 4, "OFFSET");
            FormulaNode sum = result.Root.Children[1];
            AssertNode(sum, FormulaNodeKind.Function, "SUM('O''Brien'!$A$1:$B$2)", formula.IndexOf("SUM", StringComparison.Ordinal), "SUM");
            AssertEqual(FormulaNodeKind.Binary, sum.Children[0].Kind, "qualified range is binary range");
            AssertEqual(":", sum.Children[0].Operator, "range operator");
            AssertEqual("'O''Brien'!$A$1", sum.Children[0].Children[0].Text, "left qualified span");
        }

        private static void QuotedReferences()
        {
            ReferenceArea area;
            Assert(ReferenceParser.TryParse("'[Budget.xlsx]O''Brien'!$A$1:$B$2", "DefaultBook", "DefaultSheet", out area),
                "quoted external reference parses");
            AssertEqual("Budget.xlsx", area.Workbook, "workbook is decoded");
            AssertEqual("O'Brien", area.Sheet, "quoted sheet apostrophe is decoded");
            AssertEqual(1, area.StartRow, "area first row");
            AssertEqual(2, area.EndColumn, "area last column");

            Assert(ReferenceParser.TryParse("Sheet2!A1", "Budget.xlsx", "Sheet1", out area),
                "sheet-qualified reference parses");
            AssertEqual("Budget.xlsx", area.Workbook, "sheet qualifier retains the current workbook");
            AssertEqual("Sheet2", area.Sheet, "sheet qualifier selects the target sheet");
            Assert(ReferenceParser.TryParse("'O''Brien'!$A$1:$B$2", "Budget.xlsx", "Sheet1", out area),
                "quoted sheet-qualified range parses");
            AssertEqual("Budget.xlsx", area.Workbook, "quoted sheet qualifier retains the current workbook");
            AssertEqual("O'Brien", area.Sheet, "quoted sheet qualifier is decoded");

            FormulaParseResult result = FormulaParser.Parse("=Sheet2!A1+'Quarter 1'!B2");
            Assert(result.Success, Diagnostics(result));
            AssertEqual(FormulaNodeKind.Reference, result.Root.Children[0].Kind, "unquoted sheet reference");
            AssertEqual(FormulaNodeKind.Reference, result.Root.Children[1].Kind, "quoted sheet reference");
        }

        private static void RepeatedExpressions()
        {
            FormulaParseResult result = FormulaParser.Parse("=SUM(A1+A1,A1+A1)");
            Assert(result.Success, Diagnostics(result));
            FormulaNode first = result.Root.Children[0];
            FormulaNode second = result.Root.Children[1];
            AssertEqual("A1+A1", first.Text, "first repeated expression text");
            AssertEqual("A1+A1", second.Text, "second repeated expression text");
            Assert(first.Start != second.Start, "repeated expressions retain their individual spans");
        }

        private static void MixedAbsoluteReferenceLexing()
        {
            const string formula = "=SUM(I140,$I140,I$140,$I$140,'Sheet 1'!I$140,Sheet2!$I140,Income)";
            FormulaParseResult result = FormulaParser.Parse(formula);
            Assert(result.Success, Diagnostics(result));
            AssertEqual(7, result.Root.Children.Count, "mixed absolute reference argument count");
            for (int i = 0; i < 6; i++) AssertEqual(FormulaNodeKind.Reference, result.Root.Children[i].Kind,
                "reference form " + i + " remains one reference token");
            AssertEqual("I140", result.Root.Children[0].Text, "plain reference span");
            AssertEqual("$I140", result.Root.Children[1].Text, "absolute column reference span");
            AssertEqual("I$140", result.Root.Children[2].Text, "absolute row reference span");
            AssertEqual("$I$140", result.Root.Children[3].Text, "fully absolute reference span");
            AssertEqual("'Sheet 1'!I$140", result.Root.Children[4].Text, "quoted qualified reference span");
            AssertEqual("Sheet2!$I140", result.Root.Children[5].Text, "unquoted qualified reference span");
            AssertEqual(FormulaNodeKind.Name, result.Root.Children[6].Kind, "name remains a name token");
            AssertEqual("Income", result.Root.Children[6].Text, "name span remains intact");
        }

        private static void StringAndNumericLiterals()
        {
            FormulaParseResult result = FormulaParser.Parse("=\"a\"\"b\"&1.25E-3");
            Assert(result.Success, Diagnostics(result));
            AssertEqual("&", result.Root.Operator, "concatenation operator");
            AssertEqual("\"a\"\"b\"", result.Root.Children[0].Text, "doubled quote string text");
            AssertEqual("1.25E-3", result.Root.Children[1].Text, "scientific literal text");
        }

        private static void MissingArgumentsAndArrays()
        {
            FormulaParseResult result = FormulaParser.Parse("=IF(A1,,{1,2;3,4})");
            Assert(result.Success, Diagnostics(result));
            AssertEqual(FormulaNodeKind.MissingArgument, result.Root.Children[1].Kind, "omitted IF argument");
            FormulaNode array = result.Root.Children[2];
            AssertEqual(FormulaNodeKind.Array, array.Kind, "array node");
            AssertEqual(4, array.Children.Count, "array elements");
        }

        private static void StructuredReferences()
        {
            FormulaParseResult column = FormulaParser.Parse("=Table1[Column Name]");
            Assert(column.Success, Diagnostics(column));
            AssertEqual(FormulaNodeKind.Name, column.Root.Kind, "table column is a name node");
            AssertEqual("Table1[Column Name]", column.Root.Text, "table column span preserves whitespace");
            AssertEqual(column.Root.Text, column.Root.Name, "table column name preserves source text");

            FormulaParseResult compound = FormulaParser.Parse("=Table1[[#Headers],[A]:[B]]");
            Assert(compound.Success, Diagnostics(compound));
            AssertEqual("Table1[[#Headers],[A]:[B]]", compound.Root.Text, "double bracket item/column form preserves source text");

            FormulaParseResult row = FormulaParser.Parse("=[@Col]");
            Assert(row.Success, Diagnostics(row));
            AssertEqual(FormulaNodeKind.Name, row.Root.Kind, "this-row structured reference is a name node");
            AssertEqual("[@Col]", row.Root.Text, "this-row structured reference text");

            const string expression = "=SUM(Table1[Amount],[@Col])+1";
            FormulaParseResult inExpression = FormulaParser.Parse(expression);
            Assert(inExpression.Success, Diagnostics(inExpression));
            AssertEqual("+", inExpression.Root.Operator, "structured reference does not consume the following operator");
            FormulaNode sum = inExpression.Root.Children[0];
            AssertEqual(2, sum.Children.Count, "structured reference does not consume the argument separator");
            AssertNode(sum.Children[0], FormulaNodeKind.Name, "Table1[Amount]", 5, "Table1[Amount]");
            AssertNode(sum.Children[1], FormulaNodeKind.Name, "[@Col]",
                expression.IndexOf("[@Col]", StringComparison.Ordinal), "[@Col]");
        }

        private static void MalformedFormulas()
        {
            FormulaParseResult unterminated = FormulaParser.Parse("=SUM(A1,");
            Assert(!unterminated.Success, "unterminated function cannot succeed");
            Assert(unterminated.Diagnostics.Count > 0, "unterminated function has a diagnostic");
            FormulaParseResult structured = FormulaParser.Parse("=Table1[Amount]");
            Assert(structured.Success, Diagnostics(structured));
            FormulaParseResult malformedStructured = FormulaParser.Parse("=Table1[[#Headers],[A]:[B]");
            Assert(!malformedStructured.Success, "unterminated structured reference cannot succeed");
            Assert(Contains(malformedStructured.Diagnostics, "Malformed structured reference"), "malformed structured reference diagnosis");
            FormulaParseResult missingColumnBracket = FormulaParser.Parse("=Table1[Amount");
            Assert(!missingColumnBracket.Success, "unterminated table column cannot succeed");
            Assert(Contains(missingColumnBracket.Diagnostics, "Malformed structured reference"), "unterminated table column diagnosis");
            Assert(!FormulaParser.Parse("=Table1[Amount]]").Success, "extra closing bracket cannot succeed");
        }

        private static void ParserBounds()
        {
            string nested = "=";
            for (int i = 0; i < 300; i++) nested += "(";
            nested += "A1";
            for (int i = 0; i < 300; i++) nested += ")";
            FormulaParseResult result = FormulaParser.Parse(nested);
            Assert(!result.Success, "nested formula beyond parser bound cannot succeed");
            Assert(Contains(result.Diagnostics, "maximum supported depth"), "nesting bound diagnostic");

            string additions = "=1";
            for (int i = 0; i < 3000; i++) additions += "+1";
            FormulaParseResult chain = FormulaParser.Parse(additions);
            Assert(!chain.Success, "left-associated expression beyond AST depth bound cannot succeed");
            Assert(chain.Root == null, "unsafe left-associated AST is removed");
            Assert(Contains(chain.Diagnostics, "Formula AST exceeds"), "left-associated AST depth diagnostic");
        }

        private static void ReferenceAreasAndBounds()
        {
            ReferenceArea columns;
            ReferenceArea rows;
            Assert(ReferenceParser.TryParse("$A:$XFD", null, "Sheet1", out columns), "full-column reference parses");
            AssertEqual(1, columns.StartRow, "column range begins first row");
            AssertEqual(ReferenceParser.MaxRows, columns.EndRow, "column range ends last row");
            Assert(ReferenceParser.TryParse("1:1048576", null, "Sheet1", out rows), "full-row reference parses");
            AssertEqual(1, rows.StartColumn, "row range begins first column");
            AssertEqual(ReferenceParser.MaxColumns, rows.EndColumn, "row range ends last column");
            ReferenceArea ignored;
            Assert(!ReferenceParser.TryParse("XFE1", null, null, out ignored), "column beyond XFD rejects");
            Assert(!ReferenceParser.TryParse("A1048577", null, null, out ignored), "row beyond Excel limit rejects");
            FormulaParseResult formulaRows = FormulaParser.Parse("=1:2");
            Assert(formulaRows.Success, Diagnostics(formulaRows));
            AssertEqual(FormulaNodeKind.Reference, formulaRows.Root.Children[0].Kind, "formula row-range start is a reference");
            AssertEqual(FormulaNodeKind.Reference, formulaRows.Root.Children[1].Kind, "formula row-range end is a reference");
            Assert(columns.Intersects(new ReferenceArea(null, "sheet1", 2, 4, 2, 3)), "case-insensitive same-sheet intersection");
            Assert(!columns.Intersects(new ReferenceArea(null, "sheet2", 2, 4, 2, 3)), "different sheets do not intersect");
        }

        private static void OperatorPrecedence()
        {
            FormulaParseResult result = FormulaParser.Parse("=-2^2+3*4");
            Assert(result.Success, Diagnostics(result));
            AssertEqual("+", result.Root.Operator, "addition is outermost");
            AssertEqual("^", result.Root.Children[0].Operator, "exponentiation binds after unary negation in Excel precedence");
            AssertEqual(FormulaNodeKind.Unary, result.Root.Children[0].Children[0].Kind, "negation is exponent base");
            AssertEqual("*", result.Root.Children[1].Operator, "multiplication binds before addition");
        }

        private static void DependencyPlanStaticRangesAndIntersections()
        {
            FormulaParseResult qualified = FormulaParser.Parse("=Sheet2!A1:B2");
            Assert(qualified.Success, Diagnostics(qualified));
            FormulaDependencyPlan qualifiedPlan = FormulaDependencyPlan.Create(qualified.Root, "Book.xlsx", "Sheet1");
            AssertEqual(1, qualifiedPlan.References.Count, "qualified range remains one dependency");
            AssertEqual("Sheet2", qualifiedPlan.References[0].Sheet, "left range qualifier propagates to right endpoint");
            AssertEqual(2, qualifiedPlan.References[0].EndColumn, "qualified range end column");

            FormulaParseResult expression = FormulaParser.Parse("=(A1,B2)");
            Assert(expression.Success, Diagnostics(expression));
            FormulaDependencyPlan plan = FormulaDependencyPlan.Create(expression.Root, "Book.xlsx", "Sheet1");
            AssertEqual(2, plan.References.Count, "union members are projected");
            AssertArea(plan.References[0], "Book.xlsx", "Sheet1", 1, 1, 1, 1, "union first member");
            AssertArea(plan.References[1], "Book.xlsx", "Sheet1", 2, 2, 2, 2, "union second member");
            AssertEqual(0, plan.DynamicNodes.Count, "fully static expression has no opaque nodes");

            FormulaParseResult intersection = FormulaParser.Parse("=A1:A10 A5:A15");
            Assert(intersection.Success, Diagnostics(intersection));
            FormulaDependencyPlan intersectionPlan = FormulaDependencyPlan.Create(intersection.Root, "Book.xlsx", "Sheet1");
            AssertEqual(1, intersectionPlan.References.Count, "static intersection produces one dependency");
            AssertArea(intersectionPlan.References[0], "Book.xlsx", "Sheet1", 5, 10, 1, 1, "intersection is its actual overlap");
        }

        private static void DependencyPlanDynamicAndExternalReferences()
        {
            FormulaParseResult dynamic = FormulaParser.Parse("=OFFSET(A1,B1,0)+INDIRECT(\"C1\")+A2#+Mystery(C2)");
            Assert(dynamic.Success, Diagnostics(dynamic));
            FormulaDependencyPlan plan = FormulaDependencyPlan.Create(dynamic.Root, "Book.xlsx", "Sheet1");
            AssertEqual(3, plan.References.Count, "OFFSET arguments and unknown function arguments remain dependencies");
            AssertArea(plan.References[0], "Book.xlsx", "Sheet1", 1, 1, 1, 1, "OFFSET base reference");
            AssertArea(plan.References[1], "Book.xlsx", "Sheet1", 1, 1, 2, 2, "OFFSET row argument reference");
            AssertArea(plan.References[2], "Book.xlsx", "Sheet1", 2, 2, 3, 3, "unknown function argument reference");
            AssertEqual(3, plan.DynamicNodes.Count, "OFFSET, INDIRECT and spill are opaque");
            AssertEqual("OFFSET(A1,B1,0)", plan.DynamicNodes[0].Text, "OFFSET keeps its source span");
            AssertEqual("INDIRECT(\"C1\")", plan.DynamicNodes[1].Text, "INDIRECT keeps its source span");
            AssertEqual("A2#", plan.DynamicNodes[2].Text, "spill keeps its source span");

            FormulaParseResult complexIntersection = FormulaParser.Parse("=A1 BaseName");
            Assert(complexIntersection.Success, Diagnostics(complexIntersection));
            FormulaDependencyPlan complexPlan = FormulaDependencyPlan.Create(complexIntersection.Root, "Book.xlsx", "Sheet1");
            AssertEqual(0, complexPlan.References.Count, "complex intersection does not invent an endpoint dependency");
            AssertEqual(1, complexPlan.DynamicNodes.Count, "complex intersection is opaque");
            AssertEqual("A1 BaseName", complexPlan.DynamicNodes[0].Text, "complex intersection keeps exact span");

            FormulaParseResult external = FormulaParser.Parse("=SUM([Other.xlsx]Sheet1!A1,[Other.xlsx]Sheet1!B1:C2)");
            Assert(external.Success, Diagnostics(external));
            FormulaDependencyPlan externalPlan = FormulaDependencyPlan.Create(external.Root, "Book.xlsx", "Sheet1");
            AssertEqual(0, externalPlan.References.Count, "external references are excluded from the local index");
            AssertEqual(2, externalPlan.Warnings.Count, "each external source expression reports a warning");
        }

        private static void DependencyPlanStaticOffset()
        {
            FormulaParseResult shifted = FormulaParser.Parse("=OFFSET(I$140,0,-4)+A1:A10 A5:A15");
            Assert(shifted.Success, Diagnostics(shifted));
            FormulaDependencyPlan shiftedPlan = FormulaDependencyPlan.Create(shifted.Root, "Book.xlsx", "Sheet1");
            AssertEqual(3, shiftedPlan.References.Count, "static OFFSET target, base, and intersection are retained");
            AssertArea(shiftedPlan.References[0], "Book.xlsx", "Sheet1", 140, 140, 5, 5, "shifted OFFSET target");
            AssertArea(shiftedPlan.References[1], "Book.xlsx", "Sheet1", 140, 140, 9, 9, "OFFSET base remains a dependency");
            AssertArea(shiftedPlan.References[2], "Book.xlsx", "Sheet1", 5, 10, 1, 1, "ordinary intersection still resolves");
            AssertEqual(0, shiftedPlan.DynamicNodes.Count, "literal OFFSET arguments avoid dynamic resolution");

            FormulaParseResult sized = FormulaParser.Parse("=OFFSET($H$127,0,4,2,3)");
            Assert(sized.Success, Diagnostics(sized));
            FormulaDependencyPlan sizedPlan = FormulaDependencyPlan.Create(sized.Root, "Book.xlsx", "Sheet1");
            AssertArea(sizedPlan.References[0], "Book.xlsx", "Sheet1", 127, 128, 12, 14, "literal height and width resize the target");

            FormulaParseResult rangeBase = FormulaParser.Parse("=OFFSET(A1:B2,1,1)");
            Assert(rangeBase.Success, Diagnostics(rangeBase));
            FormulaDependencyPlan rangeBasePlan = FormulaDependencyPlan.Create(rangeBase.Root, "Book.xlsx", "Sheet1");
            AssertArea(rangeBasePlan.References[0], "Book.xlsx", "Sheet1", 2, 3, 2, 3, "range OFFSET preserves its base dimensions");
            AssertArea(rangeBasePlan.References[1], "Book.xlsx", "Sheet1", 1, 2, 1, 2, "range OFFSET base remains a dependency");

            FormulaParseResult crossSheet = FormulaParser.Parse("=OFFSET(Sheet2!A1,1,2)");
            Assert(crossSheet.Success, Diagnostics(crossSheet));
            FormulaDependencyPlan crossSheetPlan = FormulaDependencyPlan.Create(crossSheet.Root, "Book.xlsx", "Sheet1");
            AssertArea(crossSheetPlan.References[0], "Book.xlsx", "Sheet2", 2, 2, 3, 3, "cross-sheet OFFSET target");
            AssertArea(crossSheetPlan.References[1], "Book.xlsx", "Sheet2", 1, 1, 1, 1, "cross-sheet OFFSET base");

            FormulaParseResult outOfBounds = FormulaParser.Parse("=OFFSET(A1,-1,0)");
            Assert(outOfBounds.Success, Diagnostics(outOfBounds));
            FormulaDependencyPlan outOfBoundsPlan = FormulaDependencyPlan.Create(outOfBounds.Root, "Book.xlsx", "Sheet1");
            AssertEqual(1, outOfBoundsPlan.DynamicNodes.Count, "out-of-bounds OFFSET remains dynamic");
            AssertEqual("OFFSET(A1,-1,0)", outOfBoundsPlan.DynamicNodes[0].Text, "dynamic OFFSET keeps its source span");
            AssertEqual(1, outOfBoundsPlan.References.Count, "out-of-bounds OFFSET still retains its base");
        }

        private static void DependencyPlanIntersectionResolutionCap()
        {
            string repeatedUnion = RepeatedUnion("A1", 50);
            FormulaParseResult pathological = FormulaParser.Parse("=" + repeatedUnion + " " + repeatedUnion);
            Assert(pathological.Success, Diagnostics(pathological));
            FormulaDependencyPlan plan = FormulaDependencyPlan.Create(pathological.Root, "Book.xlsx", "Sheet1");
            AssertEqual(0, plan.References.Count, "capped intersection skips its partial static result");
            AssertEqual(0, plan.DynamicNodes.Count, "capped intersection does not schedule an expensive dynamic resolution");
            AssertEqual(1, plan.Warnings.Count, "capped intersection has one explicit warning");
            Assert(Contains(plan.Warnings, "capped"), "cap warning explains the partial result");
        }

        private static string RepeatedUnion(string value, int count)
        {
            string result = "(";
            for (int i = 0; i < count; i++)
            {
                if (i != 0) result += ",";
                result += value;
            }
            return result + ")";
        }

        private static void DependencyRangeIndexTests()
        {
            var dependencies = new List<IndexedDependency>
            {
                new IndexedDependency(new ReferenceArea("Book.xlsx", "Sheet1", 1, ReferenceParser.MaxRows, 1, 1), 10),
                new IndexedDependency(new ReferenceArea("Book.xlsx", "Sheet1", 2, 2, 1, ReferenceParser.MaxColumns), 11),
                new IndexedDependency(new ReferenceArea("Book.xlsx", "Sheet1", 3, 4, 3, 4), 12),
                new IndexedDependency(new ReferenceArea("book.xlsx", "sheet1", 3, 3, 3, 3), 12),
                new IndexedDependency(new ReferenceArea("Book.xlsx", "Sheet2", 1, 1, 1, 1), 99)
            };
            var index = new DependencyRangeIndex(dependencies);
            AssertIds(new[] { 11 }, index.Find(new ReferenceArea("BOOK.XLSX", "SHEET1", 2, 2, 3, 3)), "whole row query");
            AssertIds(new[] { 12 }, index.Find(new ReferenceArea("Book.xlsx", "Sheet1", 3, 3, 3, 3)), "formula ids are deduplicated");
            AssertIds(new[] { 11, 12 }, index.Find(new ReferenceArea("Book.xlsx", "Sheet1", 2, 3, 3, 3)), "range query includes both rows");
            AssertIds(new[] { 10, 11 }, index.Find(new ReferenceArea("Book.xlsx", "Sheet1", 2, 2, 1, 1)), "whole column and row query");
            AssertIds(new int[0], index.Find(new ReferenceArea("Book.xlsx", "Sheet2", 2, 2, 2, 2)), "sheet groups are isolated");

            ReferenceArea[] queries =
            {
                new ReferenceArea("Book.xlsx", "Sheet1", 1, 1, 1, 1),
                new ReferenceArea("Book.xlsx", "Sheet1", 2, 2, 3, 3),
                new ReferenceArea("Book.xlsx", "Sheet1", 3, 3, 3, 3),
                new ReferenceArea("Book.xlsx", "Sheet1", 100, 100, 5, 5),
                new ReferenceArea("Book.xlsx", "Sheet2", 1, 1, 1, 1)
            };
            for (int i = 0; i < queries.Length; i++)
                AssertIds(BruteForce(dependencies, queries[i]), index.Find(queries[i]), "range index agrees with brute force query " + i);

            // Exceed the 16-entry leaf size so these checks exercise tree splitting and pruning,
            // including nested/overlapping rectangles, duplicate ids and whole rows/columns.
            var random = new Random(71239);
            for (int i = 0; i < 2048; i++)
            {
                int row = random.Next(1, 2000);
                int column = random.Next(1, 200);
                int lastRow = row + random.Next(0, 100);
                int lastColumn = column + random.Next(0, 30);
                if (i % 31 == 0) { row = 1; lastRow = ReferenceParser.MaxRows; }
                if (i % 37 == 0) { column = 1; lastColumn = ReferenceParser.MaxColumns; }
                dependencies.Add(new IndexedDependency(new ReferenceArea(i % 5 == 0 ? "Other.xlsx" : "Book.xlsx",
                    i % 3 == 0 ? "Sheet2" : "Sheet1", row, lastRow, column, lastColumn), i % 511));
            }
            index = new DependencyRangeIndex(dependencies);
            for (int i = 0; i < 500; i++)
            {
                int row = random.Next(1, 2100);
                int column = random.Next(1, 240);
                var query = new ReferenceArea(i % 5 == 0 ? "OTHER.xlsx" : "BOOK.xlsx",
                    i % 3 == 0 ? "SHEET2" : "SHEET1", row, row + random.Next(0, 25), column, column + random.Next(0, 5));
                AssertIds(BruteForce(dependencies, query), index.Find(query), "tree agrees with exhaustive intersection " + i);
            }
            var fullSheet = new ReferenceArea("Book.xlsx", "Sheet1", 1, ReferenceParser.MaxRows, 1, ReferenceParser.MaxColumns);
            AssertIds(BruteForce(dependencies, fullSheet), index.Find(fullSheet), "full-sheet query");
            AssertIds(new int[0], index.Find(new ReferenceArea("Missing.xlsx", "Sheet1", 1, 1, 1, 1)), "unknown workbook");
        }

        private static void FormulaReadBlockCoverage()
        {
            var blocks = new List<FormulaReadBlock>(FormulaReadBlocks.Enumerate(100, 129));
            AssertEqual(8, blocks.Count, "100 by 129 is split into fixed row bands across both column strips");
            bool[,] covered = new bool[101, 130];
            for (int i = 0; i < blocks.Count; i++)
            {
                FormulaReadBlock block = blocks[i];
                Assert(block.Row >= 1 && block.Column >= 1, "block offsets are one-based");
                Assert(block.Height * block.Width <= 4096, "block stays within cell limit");
                for (int row = block.Row; row < block.Row + block.Height; row++)
                    for (int column = block.Column; column < block.Column + block.Width; column++)
                    {
                        Assert(!covered[row, column], "blocks do not overlap");
                        covered[row, column] = true;
                    }
            }
            for (int row = 1; row <= 100; row++)
                for (int column = 1; column <= 129; column++)
                    Assert(covered[row, column], "every requested cell is covered");

            AssertEqual(1, new List<FormulaReadBlock>(FormulaReadBlocks.Enumerate(1, 128)).Count,
                "max-width boundary uses one strip");
            AssertEqual(2, new List<FormulaReadBlock>(FormulaReadBlocks.Enumerate(1, 129)).Count,
                "one column beyond the boundary uses two strips");
        }

        private static void AssertArea(ReferenceArea area, string workbook, string sheet, int startRow, int endRow,
            int startColumn, int endColumn, string message)
        {
            AssertEqual(workbook, area.Workbook, message + " workbook");
            AssertEqual(sheet, area.Sheet, message + " sheet");
            AssertEqual(startRow, area.StartRow, message + " start row");
            AssertEqual(endRow, area.EndRow, message + " end row");
            AssertEqual(startColumn, area.StartColumn, message + " start column");
            AssertEqual(endColumn, area.EndColumn, message + " end column");
        }

        private static IEnumerable<int> BruteForce(IList<IndexedDependency> dependencies, ReferenceArea source)
        {
            var ids = new List<int>();
            for (int i = 0; i < dependencies.Count; i++)
            {
                if (dependencies[i].Area.Intersects(source) && !ids.Contains(dependencies[i].FormulaId))
                    ids.Add(dependencies[i].FormulaId);
            }
            ids.Sort();
            return ids;
        }

        private static void AssertIds(IEnumerable<int> expected, IReadOnlyList<int> actual, string message)
        {
            var expectedValues = new List<int>(expected);
            AssertEqual(expectedValues.Count, actual.Count, message + " count");
            for (int i = 0; i < expectedValues.Count; i++) AssertEqual(expectedValues[i], actual[i], message + " item " + i);
        }

        private static void AssertNode(FormulaNode node, FormulaNodeKind kind, string text, int start, string name)
        {
            AssertEqual(kind, node.Kind, "node kind");
            AssertEqual(text, node.Text, "node text");
            AssertEqual(start, node.Start, "node start");
            AssertEqual(text.Length, node.Length, "node length");
            AssertEqual(name, node.Name, "node name");
        }

        private static void NavigationCellLeaves()
        {
            const string formula = "=IF(A1>5,B2+10,\"none\")";
            var parsed = FormulaParser.Parse(formula);
            Assert(parsed.Success, Diagnostics(parsed));
            var tree = FormulaNavigationNode.Create(parsed.Root);
            AssertEqual("IF", tree.Expression.Name, "IF remains an operation");
            AssertEqual(2, tree.Children.Count, "literal IF branch is omitted from the view");
            var leaves = NavigationLeaves(tree, formula);
            AssertEqual(2, leaves.Count, "only the two cell references are terminal nodes");
            AssertEqual("A1", leaves[0].Expression.Text, "condition cell");
            AssertEqual("B2", leaves[1].Expression.Text, "true branch cell");
            AssertEqual(3, parsed.Root.Children.Count, "original IF still has all arguments for evaluation");
            AssertEqual("\"none\"", parsed.Root.Children[2].Text, "original string is preserved");

            var lastArgument = FormulaNavigationNode.Create(FormulaParser.Parse("=IF(TRUE,0,C3)").Root);
            AssertEqual(1, lastArgument.Children.Count, "only the false branch refers to a cell");
            AssertEqual(2, lastArgument.Children[0].ArgumentIndex, "omission does not relabel value_if_false as logical_test");
            var missing = FormulaNavigationNode.Create(FormulaParser.Parse("=IF(A1,,D2)").Root);
            AssertEqual(2, missing.Children[1].ArgumentIndex, "omitted argument does not shift later argument labels");

            Assert(FormulaNavigationNode.Create(FormulaParser.Parse("=SUM(1,2)*3").Root) == null,
                "reference-free arithmetic has no fake terminal operation");
            Assert(FormulaNavigationNode.Create(FormulaParser.Parse("=IF(FALSE,{1,2},\"x\")").Root) == null,
                "booleans, arrays and strings do not create leaves");
            Assert(FormulaNavigationNode.Create(FormulaParser.Parse("=TODAY()").Root) == null,
                "function with no cell references stays only in source text");
        }

        private static void NavigationRangesAndSpans()
        {
            const string formula = "=SUM('Sheet 1'!A1:A4,A1+A1,Table1[Amount],BaseCell,A2#)";
            var parsed = FormulaParser.Parse(formula);
            Assert(parsed.Success, Diagnostics(parsed));
            var leaves = NavigationLeaves(FormulaNavigationNode.Create(parsed.Root), formula);
            AssertEqual(6, leaves.Count, "ranges, repeated cells, names and spill references are retained");
            AssertEqual("'Sheet 1'!A1:A4", leaves[0].Expression.Text, "range is one complete reference");
            AssertEqual("A1", leaves[1].Expression.Text, "first occurrence");
            AssertEqual("A1", leaves[2].Expression.Text, "second occurrence");
            Assert(leaves[1].Expression.Start != leaves[2].Expression.Start, "repeated references keep independent highlights");
            AssertEqual("Table1[Amount]", leaves[3].Expression.Text, "structured reference remains");
            AssertEqual("BaseCell", leaves[4].Expression.Text, "named reference remains");
            AssertEqual("A2#", leaves[5].Expression.Text, "spill is one complete reference");
        }

        private static void NavigationCalculatedReferences()
        {
            const string formula = "=IF(1,OFFSET(A1,B1,0),INDIRECT(\"C\"&2))";
            var parsed = FormulaParser.Parse(formula);
            Assert(parsed.Success, Diagnostics(parsed));
            var tree = FormulaNavigationNode.Create(parsed.Root);
            var leaves = NavigationLeaves(tree, formula);
            AssertEqual(4, leaves.Count, "reference arguments and both calculated destinations remain");
            AssertEqual("A1", leaves[0].Expression.Text, "OFFSET base reference");
            AssertEqual("B1", leaves[1].Expression.Text, "cell controlling the offset");
            AssertEqual(FormulaNavigationNodeKind.CalculatedReference, leaves[2].Kind, "OFFSET result is a reference leaf");
            AssertEqual("OFFSET(A1,B1,0)", leaves[2].Expression.Text, "OFFSET destination highlights its full expression");
            AssertEqual(FormulaNavigationNodeKind.CalculatedReference, leaves[3].Kind, "INDIRECT result is not pruned with literal arguments");
            AssertEqual("INDIRECT(\"C\"&2)", leaves[3].Expression.Text, "INDIRECT destination retains exact original span");
            AssertEqual(1, tree.Children[0].ArgumentIndex, "OFFSET remains value_if_true");
            AssertEqual(2, tree.Children[1].ArgumentIndex, "INDIRECT remains value_if_false");
        }

        private static List<FormulaNavigationNode> NavigationLeaves(FormulaNavigationNode tree, string formula)
        {
            var leaves = new List<FormulaNavigationNode>();
            CollectNavigationLeaves(tree, formula, leaves);
            return leaves;
        }

        private static void CollectNavigationLeaves(FormulaNavigationNode node, string formula, List<FormulaNavigationNode> leaves)
        {
            AssertEqual(node.Expression.Text, formula.Substring(node.Expression.Start, node.Expression.Length),
                "navigation projection preserves source offsets");
            if (node.Children.Count == 0)
            {
                Assert(node.Kind != FormulaNavigationNodeKind.Operation, "terminal nodes must be references");
                leaves.Add(node);
            }
            else foreach (var child in node.Children) CollectNavigationLeaves(child, formula, leaves);
        }

        private static bool Contains(System.Collections.Generic.IReadOnlyList<string> values, string part)
        {
            for (int i = 0; i < values.Count; i++) if (values[i].IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        private static string Diagnostics(FormulaParseResult result)
        {
            return result.Diagnostics.Count == 0 ? "no diagnostics" : String.Join(" | ", result.Diagnostics);
        }

        private static void Run(string name, Action action)
        {
            try { action(); Console.WriteLine("PASS " + name); }
            catch (Exception exception) { _failures++; Console.Error.WriteLine("FAIL " + name + ": " + exception.Message); }
        }

        private static void Assert(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
        private static void AssertEqual<T>(T expected, T actual, string message)
        {
            if (!Object.Equals(expected, actual)) throw new InvalidOperationException(message + "; expected '" + expected + "', actual '" + actual + "'.");
        }
    }

    // Small, dependency-free PE32+ verifier used by the Linux build pipeline after packing the XLL.
    internal static class XllVerifier
    {
        private const ushort MachineAmd64 = 0x8664;
        private const ushort DllCharacteristic = 0x2000;

        public static bool TryVerify(string path, out string error)
        {
            error = null;
            byte[] image;
            try { image = File.ReadAllBytes(path); }
            catch (Exception exception) { error = "cannot read file: " + exception.Message; return false; }
            if (!Has(image, 0, 64) || image[0] != 'M' || image[1] != 'Z') { error = "missing MZ header"; return false; }
            uint peOffsetValue = U32(image, 0x3c);
            if (peOffsetValue > (uint)Int32.MaxValue) { error = "invalid PE offset"; return false; }
            int peOffset = (int)peOffsetValue;
            if (!Has(image, peOffset, 24) || U32(image, peOffset) != 0x00004550u) { error = "missing PE signature"; return false; }
            if (U16(image, peOffset + 4) != MachineAmd64) { error = "PE machine is not AMD64"; return false; }
            if ((U16(image, peOffset + 22) & DllCharacteristic) == 0) { error = "PE is not marked as a DLL"; return false; }

            int optional = peOffset + 24;
            int optionalSize = U16(image, peOffset + 20);
            if (!Has(image, optional, optionalSize) || optionalSize < 120 || U16(image, optional) != 0x20b)
            {
                error = "not a PE32+ optional header"; return false;
            }
            uint exportRva = U32(image, optional + 112);
            if (exportRva == 0) { error = "PE has no export table"; return false; }
            int sections = U16(image, peOffset + 6);
            int sectionOffset = optional + optionalSize;
            if (!Has(image, sectionOffset, sections * 40)) { error = "truncated section table"; return false; }
            int exportOffset;
            if (!RvaToOffset(image, optional, sectionOffset, sections, exportRva, out exportOffset) || !Has(image, exportOffset, 40))
            {
                error = "export directory is outside the image"; return false;
            }
            uint numberOfNames = U32(image, exportOffset + 24);
            uint namesRva = U32(image, exportOffset + 32);
            if (numberOfNames == 0 || numberOfNames > 65536 || namesRva == 0) { error = "export table has no usable named exports"; return false; }
            int namesOffset;
            if (!RvaToOffset(image, optional, sectionOffset, sections, namesRva, out namesOffset)
                || !Has(image, namesOffset, checked((int)numberOfNames * 4)))
            {
                error = "export name pointer table is invalid"; return false;
            }
            for (int i = 0; i < (int)numberOfNames; i++)
            {
                int nameOffset;
                if (!RvaToOffset(image, optional, sectionOffset, sections, U32(image, namesOffset + i * 4), out nameOffset)) continue;
                string name;
                if (AsciiZ(image, nameOffset, out name) && name == "xlAutoOpen") return true;
            }
            error = "named export xlAutoOpen was not found";
            return false;
        }

        private static bool RvaToOffset(byte[] image, int optionalOffset, int sectionOffset, int sections, uint rva, out int offset)
        {
            offset = 0;
            uint headersSize = U32(image, optionalOffset + 60);
            if (rva < headersSize && rva < (uint)image.Length) { offset = (int)rva; return true; }
            for (int i = 0; i < sections; i++)
            {
                int section = sectionOffset + i * 40;
                uint virtualSize = U32(image, section + 8);
                uint virtualAddress = U32(image, section + 12);
                uint rawSize = U32(image, section + 16);
                uint rawOffset = U32(image, section + 20);
                uint size = Math.Max(virtualSize, rawSize);
                if (rva < virtualAddress || rva - virtualAddress >= size) continue;
                ulong candidate = (ulong)rawOffset + (rva - virtualAddress);
                if (candidate >= (ulong)image.Length) return false;
                offset = (int)candidate;
                return true;
            }
            return false;
        }

        private static bool AsciiZ(byte[] image, int offset, out string value)
        {
            value = null;
            if (offset < 0 || offset >= image.Length) return false;
            int end = offset;
            while (end < image.Length && end - offset < 512 && image[end] != 0) end++;
            if (end == image.Length || end - offset == 512) return false;
            char[] chars = new char[end - offset];
            for (int i = 0; i < chars.Length; i++) chars[i] = (char)image[offset + i];
            value = new String(chars);
            return true;
        }

        private static bool Has(byte[] image, int offset, int count)
        {
            return offset >= 0 && count >= 0 && offset <= image.Length && count <= image.Length - offset;
        }

        private static ushort U16(byte[] image, int offset) { return (ushort)(image[offset] | image[offset + 1] << 8); }
        private static uint U32(byte[] image, int offset)
        {
            return (uint)(image[offset] | image[offset + 1] << 8 | image[offset + 2] << 16 | image[offset + 3] << 24);
        }
    }
}
