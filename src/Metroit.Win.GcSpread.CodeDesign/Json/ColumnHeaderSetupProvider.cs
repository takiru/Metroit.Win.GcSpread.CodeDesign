using FarPoint.Win.Spread;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// ColumnHeader のセットアップ操作を提供します。
    /// </summary>
    public class ColumnHeaderSetupProvider
    {
        /// <summary>
        /// 列ヘッダーセルに自動付与される Tag の接頭辞。
        /// </summary>
        private static readonly string HeaderCellTagPrefix = @"Header";

        /// <summary>
        /// 処理の対象となる SheetView。
        /// </summary>
        private SheetView SheetView { get; }

        /// <summary>
        /// 新しい ColumnHeaderSetupProvider インスタンスを生成します。
        /// </summary>
        /// <param name="sheetView">処理対象とする SheetView。</param>
        public ColumnHeaderSetupProvider(SheetView sheetView)
        {
            SheetView = sheetView;
        }

        /// <summary>
        /// 列ヘッダー行定義から列ヘッダー行をセットアップします。
        /// 行は必ず1行目から追加されます。
        /// </summary>
        /// <param name="headerRowDefs">列ヘッダー行定義オブジェクト。</param>
        /// <param name="templateLayoutDefsList">テンプレート定義オブジェクト。</param>
        public void SetupRows(HeaderRowDefinitions[] headerRowDefs, List<TemplateSheetViewDefinitions> templateLayoutDefsList = null)
        {
            if (headerRowDefs == null)
            {
                return;
            }

            SheetView.ColumnHeader.Rows.Add(0, headerRowDefs.Length);
            foreach (var row in headerRowDefs.Select((Item, Index) => new { Item, Index }))
            {
                // 列ヘッダー行テンプレートから定義を反映する
                if (templateLayoutDefsList != null)
                {
                    var targetTemplates = new List<TemplateHeaderRowDefinitions>();
                    ReadRowTemplates(templateLayoutDefsList, row.Item, targetTemplates);
                    foreach (var targetTemplate in targetTemplates)
                    {
                        SetColumnHeaderRowProps(SheetView.ColumnHeader.Rows[row.Index], targetTemplate);
                    }
                }

                SetColumnHeaderRowProps(SheetView.ColumnHeader.Rows[row.Index], row.Item);
            }
        }

        /// <summary>
        /// 列ヘッダーセル定義から列ヘッダーセルをセットアップします。
        /// 列は必ず1列目から追加されます。
        /// </summary>
        /// <param name="headerCellDefs">列ヘッダーセル定義オブジェクト。</param>
        /// <param name="cellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        public void SetupCells(HeaderCellDefinitions[][] headerCellDefs, Func<Cell, object> cellTag = null)
        {
            SetupCells(headerCellDefs, null, cellTag);
        }

        /// <summary>
        /// 列ヘッダーセル定義から列ヘッダーセルをセットアップします。
        /// 列は必ず1列目から追加されます。
        /// </summary>
        /// <param name="headerCellDefs">列ヘッダーセル定義オブジェクト。</param>
        /// <param name="templateLayoutDefsList">テンプレート定義オブジェクト。</param>
        /// <param name="cellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        public void SetupCells(HeaderCellDefinitions[][] headerCellDefs, List<TemplateSheetViewDefinitions> templateLayoutDefsList, Func<Cell, object> cellTag = null)
        {
            if (headerCellDefs == null)
            {
                return;
            }

            SheetView.ColumnHeader.Columns.Add(0, headerCellDefs.Max(x => x.Length));
            foreach (var cells in headerCellDefs.Select((Item, Index) => new { Item, Index }))
            {
                foreach (var cell in cells.Item.Select((Item, Index) => new { Item, Index }))
                {
                    // 列ヘッダーセルテンプレートから定義を反映する
                    if (templateLayoutDefsList != null)
                    {
                        var targetTemplates = new List<TemplateHeaderCellDefinitions>();
                        ReadCellTemplates(templateLayoutDefsList, cell.Item, targetTemplates);
                        foreach (var targetTemplate in targetTemplates)
                        {
                            SetColumnHeaderCellProps(SheetView.ColumnHeader.Cells[cells.Index, cell.Index], targetTemplate);
                        }
                    }

                    SetColumnHeaderCellProps(SheetView.ColumnHeader.Cells[cells.Index, cell.Index], cell.Item);

                    // Tag の設定
                    if (cellTag == null)
                    {
                        SheetView.ColumnHeader.Cells[cells.Index, cell.Index].Tag = $"{HeaderCellTagPrefix}_{cells.Index}_{cell.Index}";
                    }
                    else
                    {
                        SheetView.ColumnHeader.Cells[cells.Index, cell.Index].Tag = cellTag.Invoke(SheetView.ColumnHeader.Cells[cells.Index, cell.Index]);
                    }
                }
            }
        }

        /// <summary>
        /// 列ヘッダースパン定義から列ヘッダーセルスパンを生成します。
        /// </summary>
        /// <param name="spanDefs">列ヘッダースパン定義オブジェクト。</param>
        /// <param name="columnMoveResults">列移動によって spans の定義情報と合致しない列情報。</param>
        public void SetupSpans(SpanDefinitions[] spanDefs, IList<ColumnMoveResult> columnMoveResults = null)
        {
            if (spanDefs == null)
            {
                return;
            }

            SpanDefinitions[] newSpans;

            if (columnMoveResults == null)
            {
                newSpans = spanDefs;
            }
            else
            {
                // 列移動情報がある場合は、オリジナルの列インデックスでマッピングして、移動後のインデックスでセル結合を行う
                var spanList = new List<SpanDefinitions>();
                foreach (var span in spanDefs)
                {
                    var columnMoveResult = columnMoveResults.Where(x => x.OriginalColumnIndex == span.Column).FirstOrDefault();
                    if (columnMoveResult == null)
                    {
                        spanList.Add(span);
                    }
                    else
                    {
                        spanList.Add(new SpanDefinitions()
                        {
                            Row = span.Row,
                            Column = columnMoveResult.AfterColumnIndex,
                            RowCount = span.RowCount,
                            ColumnCount = span.ColumnCount
                        });
                    }
                }
                newSpans = spanList.ToArray();
            }

            foreach (var span in newSpans)
            {
                SheetView.AddColumnHeaderSpanCell(span.Row, span.Column, span.RowCount, span.ColumnCount);
            }
        }

        /// <summary>
        /// 列ヘッダー行定義を生成します。
        /// </summary>
        /// <param name="includeProps">生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>HeaderRowDefinitions[] オブジェクト。</returns>
        public HeaderRowDefinitions[] CreateRowDefinitions(string[] includeProps = null)
        {
            var headerRows = new List<HeaderRowDefinitions>();
            foreach (Row row in SheetView.ColumnHeader.Rows)
            {
                var headerRow = new HeaderRowDefinitions();

                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Row.Height))))
                {
                    headerRow.Height = row.Height;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Row.HorizontalAlignment))))
                {
                    headerRow.HorizontalAlignment = row.HorizontalAlignment;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Row.VerticalAlignment))))
                {
                    headerRow.VerticalAlignment = row.VerticalAlignment;
                }

                headerRows.Add(headerRow);
            }

            return headerRows.ToArray();
        }

        /// <summary>
        /// 列ヘッダーセル定義を生成します。
        /// </summary>
        /// <param name="includeProps">生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>HeaderCellDefinitions[][] オブジェクト。</returns>
        public HeaderCellDefinitions[][] CreateCellDefinitions(string[] includeProps = null)
        {
            var allCells = new List<HeaderCellDefinitions[]>();
            foreach (Row row in SheetView.ColumnHeader.Rows)
            {
                var cells = new List<HeaderCellDefinitions>();
                foreach (Column column in SheetView.ColumnHeader.Columns)
                {
                    var cell = new HeaderCellDefinitions();

                    if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Cell.Value))))
                    {
                        cell.Value = SheetView.ColumnHeader.Cells[row.Index, column.Index].Value?.ToString();
                    }
                    if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Cell.HorizontalAlignment))))
                    {
                        cell.HorizontalAlignment = SheetView.ColumnHeader.Cells[row.Index, column.Index].HorizontalAlignment;
                    }
                    if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Cell.VerticalAlignment))))
                    {
                        cell.VerticalAlignment = SheetView.ColumnHeader.Cells[row.Index, column.Index].VerticalAlignment;
                    }

                    cells.Add(cell);
                }
                allCells.Add(cells.ToArray());
            }

            return allCells.ToArray();
        }

        /// <summary>
        /// 列ヘッダースパン定義を生成します。
        /// </summary>
        /// <returns>SpanDefinitions[] オブジェクト。</returns>
        public SpanDefinitions[] CreateSpanDefinitions()
        {
            var headerSpans = new List<SpanDefinitions>();
            foreach (Row row in SheetView.ColumnHeader.Rows)
            {
                foreach (Column column in SheetView.ColumnHeader.Columns)
                {
                    var spanRange = SheetView.GetColumnHeaderSpanCell(row.Index, column.Index);
                    if (spanRange == null)
                    {
                        continue;
                    }
                    // 把握済のスパン情報はスキップ
                    if (headerSpans.Any(x => x.Row == spanRange.Row &&
                            x.Column == spanRange.Column &&
                            x.RowCount == spanRange.RowCount &&
                            x.ColumnCount == spanRange.ColumnCount))
                    {
                        continue;
                    }
                    headerSpans.Add(new SpanDefinitions()
                    {
                        Row = spanRange.Row,
                        Column = spanRange.Column,
                        RowCount = spanRange.RowCount,
                        ColumnCount = spanRange.ColumnCount
                    });
                }
            }

            return headerSpans.ToArray();
        }

        /// <summary>
        /// テンプレートから対象の列ヘッダー行定義テンプレートを取得する。
        /// </summary>
        /// <param name="templateLayoutDefsList">テンプレート定義オブジェクト。</param>
        /// <param name="headerRowDefs">列ヘッダー行定義オブジェクト。</param>
        /// <param name="targetTemplates">対象となった列ヘッダー行定義オブジェクト。</param>
        private void ReadRowTemplates(List<TemplateSheetViewDefinitions> templateLayoutDefsList, HeaderRowDefinitions headerRowDefs, List<TemplateHeaderRowDefinitions> targetTemplates)
        {
            foreach (var templateLayoutDefs in templateLayoutDefsList)
            {
                var templateHeaderRowDefs = templateLayoutDefs.ColumnHeader.Rows.Where(x => x.TemplateName == headerRowDefs.BaseTemplate).FirstOrDefault();
                if (templateHeaderRowDefs == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(templateHeaderRowDefs.TemplateName))
                {
                    ReadRowTemplates(templateLayoutDefsList, templateHeaderRowDefs, targetTemplates);
                }
                targetTemplates.Add(templateHeaderRowDefs);
            }
        }

        /// <summary>
        /// テンプレートから対象の列ヘッダーセル定義テンプレートを取得する。
        /// </summary>
        /// <param name="templateLayoutDefsList">テンプレート定義オブジェクト。</param>
        /// <param name="headerCellDefs">列ヘッダーセル定義オブジェクト。</param>
        /// <param name="targetTemplates">対象となった列ヘッダーセル定義オブジェクト。</param>
        private void ReadCellTemplates(List<TemplateSheetViewDefinitions> templateLayoutDefsList, HeaderCellDefinitions headerCellDefs, List<TemplateHeaderCellDefinitions> targetTemplates)
        {
            foreach (var templateLayoutDefs in templateLayoutDefsList)
            {
                var templateHeaderCellDefs = templateLayoutDefs.ColumnHeader.Cells.Where(x => x.TemplateName == headerCellDefs.BaseTemplate).FirstOrDefault();
                if (templateHeaderCellDefs == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(templateHeaderCellDefs.TemplateName))
                {
                    ReadCellTemplates(templateLayoutDefsList, templateHeaderCellDefs, targetTemplates);
                }
                targetTemplates.Add(templateHeaderCellDefs);
            }
        }

        /// <summary>
        /// 列ヘッダー行のプロパティを設定する。
        /// </summary>
        /// <param name="row">行オブジェクト。</param>
        /// <param name="headerRowDefs">列ヘッダー行定義オブジェクト。</param>
        private static void SetColumnHeaderRowProps(Row row, HeaderRowDefinitions headerRowDefs)
        {
            if (headerRowDefs.Height.HasValue)
            {
                row.Height = headerRowDefs.Height.Value;
            }
            if (headerRowDefs.HorizontalAlignment.HasValue)
            {
                row.HorizontalAlignment = headerRowDefs.HorizontalAlignment.Value;
            }
            if (headerRowDefs.VerticalAlignment.HasValue)
            {
                row.VerticalAlignment = headerRowDefs.VerticalAlignment.Value;
            }
        }

        /// <summary>
        /// 列ヘッダーセルのプロパティを設定する。
        /// </summary>
        /// <param name="cell">セルオブジェクト。</param>
        /// <param name="headerCellDefs">列ヘッダーセル定義オブジェクト。</param>
        private static void SetColumnHeaderCellProps(Cell cell, HeaderCellDefinitions headerCellDefs)
        {
            if (!string.IsNullOrEmpty(headerCellDefs.Value))
            {
                cell.Value = headerCellDefs.Value;
            }
            if (headerCellDefs.HorizontalAlignment.HasValue)
            {
                cell.HorizontalAlignment = headerCellDefs.HorizontalAlignment.Value;
            }
            if (headerCellDefs.VerticalAlignment.HasValue)
            {
                cell.VerticalAlignment = headerCellDefs.VerticalAlignment.Value;
            }
        }
    }
}
