using FarPoint.Win.Spread;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// ColumnHeader の生成操作を提供します。
    /// </summary>
    public class ColumnHeaderGenerator
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
        /// 新しい ColumnHeaderGenerator インスタンスを生成します。
        /// </summary>
        /// <param name="sheetView">処理対象とする SheetView。</param>
        public ColumnHeaderGenerator(SheetView sheetView)
        {
            SheetView = sheetView;
        }

        /// <summary>
        /// 列ヘッダー行定義オブジェクトからJSON文字列にシリアライズします。
        /// </summary>
        /// <param name="defs">列ヘッダー行定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeRowJson(HeaderRowDefinitions[] defs)
        {
            return JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// JSON文字列から列ヘッダー行定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列ヘッダー行定義JSON文字列。</param>
        /// <returns>列ヘッダー行定義オブジェクト。</returns>
        public static HeaderRowDefinitions[] DeserializeRowJson(string json)
        {
            return JsonConvert.DeserializeObject<HeaderRowDefinitions[]>(json);
        }

        /// <summary>
        /// 列ヘッダーセル定義オブジェクトからJSON文字列にシリアライズします。
        /// </summary>
        /// <param name="defs">列ヘッダーセル定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeCellJson(HeaderCellDefinitions[][] defs)
        {
            return JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// JSON文字列から列ヘッダーセル定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列ヘッダーセル定義JSON文字列。</param>
        /// <returns>列ヘッダーセル定義オブジェクト。</returns>
        public static HeaderCellDefinitions[][] DeserializeCellJson(string json)
        {
            return JsonConvert.DeserializeObject<HeaderCellDefinitions[][]>(json);
        }

        /// <summary>
        /// 列ヘッダースパン定義オブジェクトからJSON文字列にシリアライズします。
        /// </summary>
        /// <param name="defs">列ヘッダーセル定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeSpanJson(SpanDefinitions[] defs)
        {
            return JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// JSON文字列から列ヘッダースパン定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列ヘッダースパン定義JSON文字列。</param>
        /// <returns>列ヘッダースパン定義オブジェクト。</returns>
        public static SpanDefinitions[] DeserializeSpanJson(string json)
        {
            return JsonConvert.DeserializeObject<SpanDefinitions[]>(json);
        }


        /// <summary>
        /// 列ヘッダー行定義から列ヘッダー行をセットアップします。
        /// 行は必ず1行目から追加されます。
        /// </summary>
        /// <param name="defs">列ヘッダー行定義オブジェクト。</param>
        /// <param name="templateDefs">テンプレート定義オブジェクト。</param>
        public void GenerateRows(HeaderRowDefinitions[] defs, List<TemplateSheetViewDefinitions> templateDefs = null)
        {
            if (defs == null)
            {
                return;
            }

            SheetView.ColumnHeader.Rows.Add(0, defs.Length);
            foreach (var row in defs.Select((Item, Index) => new { Item, Index }))
            {
                // 列ヘッダー行テンプレートから定義を反映する
                if (templateDefs != null)
                {
                    var targetTemplates = new List<TemplateHeaderRowDefinitions>();
                    ReadRowTemplates(templateDefs, row.Item, targetTemplates);
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
        /// <param name="defs">列ヘッダーセル定義オブジェクト。</param>
        /// <param name="cellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        public void GenerateCells(HeaderCellDefinitions[][] defs, Func<Cell, object> cellTag = null)
        {
            GenerateCells(defs, null, cellTag);
        }

        /// <summary>
        /// 列ヘッダーセル定義から列ヘッダーセルをセットアップします。
        /// 列は必ず1列目から追加されます。
        /// </summary>
        /// <param name="defs">列ヘッダーセル定義オブジェクト。</param>
        /// <param name="templateDefs">テンプレート定義オブジェクト。</param>
        /// <param name="cellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        public void GenerateCells(HeaderCellDefinitions[][] defs, List<TemplateSheetViewDefinitions> templateDefs, Func<Cell, object> cellTag = null)
        {
            if (defs == null)
            {
                return;
            }

            SheetView.ColumnHeader.Columns.Add(0, defs.Max(x => x.Length));
            foreach (var cells in defs.Select((Item, Index) => new { Item, Index }))
            {
                foreach (var cell in cells.Item.Select((Item, Index) => new { Item, Index }))
                {
                    // 列ヘッダーセルテンプレートから定義を反映する
                    if (templateDefs != null)
                    {
                        var targetTemplates = new List<TemplateHeaderCellDefinitions>();
                        ReadCellTemplates(templateDefs, cell.Item, targetTemplates);
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
        /// <param name="defs">列ヘッダースパン定義オブジェクト。</param>
        /// <param name="columnMoveResults">列移動によって defs の定義情報と合致しない列情報。</param>
        public void GenerateSpans(SpanDefinitions[] defs, IList<ColumnMoveResult> columnMoveResults = null)
        {
            if (defs == null)
            {
                return;
            }

            SpanDefinitions[] newSpanDefs;

            if (columnMoveResults == null)
            {
                newSpanDefs = defs;
            }
            else
            {
                // 列移動情報がある場合は、オリジナルの列インデックスでマッピングして、移動後のインデックスでセル結合を行う
                var spanList = new List<SpanDefinitions>();
                foreach (var span in defs)
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
                newSpanDefs = spanList.ToArray();
            }

            foreach (var spanDef in newSpanDefs)
            {
                SheetView.AddColumnHeaderSpanCell(spanDef.Row, spanDef.Column, spanDef.RowCount, spanDef.ColumnCount);
            }
        }

        /// <summary>
        /// 列ヘッダー行定義を生成します。
        /// </summary>
        /// <param name="includeProps">生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>列ヘッダー行定義オブジェクト。</returns>
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
        /// <returns>列ヘッダーセル定義オブジェクト。</returns>
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
        /// <returns>列ヘッダースパン定義オブジェクト。</returns>
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
        /// テンプレートから対象のテンプレート列ヘッダー行定義オブジェクトを取得する。
        /// </summary>
        /// <param name="templateDefs">テンプレート定義オブジェクト。</param>
        /// <param name="def">列ヘッダー行定義オブジェクト。</param>
        /// <param name="targetTemplateDefs">対象となったテンプレート列ヘッダー行定義オブジェクト。</param>
        private void ReadRowTemplates(List<TemplateSheetViewDefinitions> templateDefs, HeaderRowDefinitions def, List<TemplateHeaderRowDefinitions> targetTemplateDefs)
        {
            foreach (var templateDef in templateDefs)
            {
                if (templateDef.ColumnHeader == null)
                {
                    continue;
                }
                if (templateDef.ColumnHeader.Rows == null)
                {
                    continue;
                }

                var rowDef = templateDef.ColumnHeader.Rows.Where(x => x.TemplateName == def.BaseTemplate).FirstOrDefault();
                if (rowDef == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(rowDef.TemplateName))
                {
                    ReadRowTemplates(templateDefs, rowDef, targetTemplateDefs);
                }
                targetTemplateDefs.Add(rowDef);
            }
        }

        /// <summary>
        /// テンプレートから対象のテンプレート列ヘッダーセル定義オブジェクトを取得する。
        /// </summary>
        /// <param name="templateDefs">テンプレート定義オブジェクト。</param>
        /// <param name="def">列ヘッダーセル定義オブジェクト。</param>
        /// <param name="targetTemplateDefs">対象となったテンプレート列ヘッダーセル定義オブジェクト。</param>
        private void ReadCellTemplates(List<TemplateSheetViewDefinitions> templateDefs, HeaderCellDefinitions def, List<TemplateHeaderCellDefinitions> targetTemplateDefs)
        {
            foreach (var templateDef in templateDefs)
            {
                if (templateDef.ColumnHeader == null)
                {
                    continue;
                }
                if (templateDef.ColumnHeader.Cells == null)
                {
                    continue;
                }

                var cellDef = templateDef.ColumnHeader.Cells.Where(x => x.TemplateName == def.BaseTemplate).FirstOrDefault();
                if (cellDef == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(cellDef.TemplateName))
                {
                    ReadCellTemplates(templateDefs, cellDef, targetTemplateDefs);
                }
                targetTemplateDefs.Add(cellDef);
            }
        }

        /// <summary>
        /// 列ヘッダー行のプロパティを設定する。
        /// </summary>
        /// <param name="row">行オブジェクト。</param>
        /// <param name="def">列ヘッダー行定義オブジェクト。</param>
        private static void SetColumnHeaderRowProps(Row row, HeaderRowDefinitions def)
        {
            if (def.Height.HasValue)
            {
                row.Height = def.Height.Value;
            }
            if (def.HorizontalAlignment.HasValue)
            {
                row.HorizontalAlignment = def.HorizontalAlignment.Value;
            }
            if (def.VerticalAlignment.HasValue)
            {
                row.VerticalAlignment = def.VerticalAlignment.Value;
            }
        }

        /// <summary>
        /// 列ヘッダーセルのプロパティを設定する。
        /// </summary>
        /// <param name="cell">セルオブジェクト。</param>
        /// <param name="def">列ヘッダーセル定義オブジェクト。</param>
        private static void SetColumnHeaderCellProps(Cell cell, HeaderCellDefinitions def)
        {
            if (!string.IsNullOrEmpty(def.Value))
            {
                cell.Value = def.Value;
            }
            if (def.HorizontalAlignment.HasValue)
            {
                cell.HorizontalAlignment = def.HorizontalAlignment.Value;
            }
            if (def.VerticalAlignment.HasValue)
            {
                cell.VerticalAlignment = def.VerticalAlignment.Value;
            }
        }
    }
}
