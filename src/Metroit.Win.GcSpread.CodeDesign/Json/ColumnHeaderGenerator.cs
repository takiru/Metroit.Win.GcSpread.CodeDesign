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
        public static string SerializeRowJson(HeaderRowDefinition[] defs)
        {
            return JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// JSON文字列から列ヘッダー行定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列ヘッダー行定義JSON文字列。</param>
        /// <returns>列ヘッダー行定義オブジェクト。</returns>
        public static HeaderRowDefinition[] DeserializeRowJson(string json)
        {
            return JsonConvert.DeserializeObject<HeaderRowDefinition[]>(json);
        }

        /// <summary>
        /// 列ヘッダーセル定義オブジェクトからJSON文字列にシリアライズします。
        /// </summary>
        /// <param name="defs">列ヘッダーセル定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeCellJson(HeaderCellDefinition[][] defs)
        {
            return JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// JSON文字列から列ヘッダーセル定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列ヘッダーセル定義JSON文字列。</param>
        /// <returns>列ヘッダーセル定義オブジェクト。</returns>
        public static HeaderCellDefinition[][] DeserializeCellJson(string json)
        {
            return JsonConvert.DeserializeObject<HeaderCellDefinition[][]>(json);
        }

        /// <summary>
        /// 列ヘッダースパン定義オブジェクトからJSON文字列にシリアライズします。
        /// </summary>
        /// <param name="defs">列ヘッダーセル定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeSpanJson(SpanDefinition[] defs)
        {
            return JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// JSON文字列から列ヘッダースパン定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列ヘッダースパン定義JSON文字列。</param>
        /// <returns>列ヘッダースパン定義オブジェクト。</returns>
        public static SpanDefinition[] DeserializeSpanJson(string json)
        {
            return JsonConvert.DeserializeObject<SpanDefinition[]>(json);
        }


        /// <summary>
        /// 列ヘッダー行定義から列ヘッダー行をセットアップします。
        /// 行は必ず1行目から追加されます。
        /// </summary>
        /// <param name="defs">列ヘッダー行定義オブジェクト。</param>
        /// <param name="templateDefs">テンプレート定義オブジェクト。</param>
        public void GenerateRows(HeaderRowDefinition[] defs, List<TemplateSheetViewDefinition> templateDefs = null)
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
                    var targetTemplates = new List<TemplateHeaderRowDefinition>();
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
        public void GenerateCells(HeaderCellDefinition[][] defs, Func<Cell, object> cellTag = null)
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
        public void GenerateCells(HeaderCellDefinition[][] defs, List<TemplateSheetViewDefinition> templateDefs, Func<Cell, object> cellTag = null)
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
                        var targetTemplates = new List<TemplateHeaderCellDefinition>();
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
        /// <param name="spanDefinitions">列ヘッダースパン定義オブジェクト。</param>
        /// <param name="columnMoveResults">列移動によって spanDefinitions の定義情報と合致しない列情報。</param>
        public void GenerateSpans(SpanDefinition[] spanDefinitions, IList<ColumnMoveResult> columnMoveResults = null)
        {
            if (spanDefinitions == null)
            {
                return;
            }

            SpanDefinition[] newSpanDefinitions;

            if (columnMoveResults == null)
            {
                newSpanDefinitions = spanDefinitions;
            }
            else
            {
                // 列移動情報がある場合は、オリジナルの列インデックスでマッピングして、移動後のインデックスでセル結合を行う
                var spanList = new List<SpanDefinition>();
                foreach (var span in spanDefinitions)
                {
                    var columnMoveResult = columnMoveResults.Where(x => x.OriginalColumnIndex == span.Column).FirstOrDefault();
                    if (columnMoveResult == null)
                    {
                        spanList.Add(span);
                    }
                    else
                    {
                        spanList.Add(new SpanDefinition()
                        {
                            Row = span.Row,
                            Column = columnMoveResult.AfterColumnIndex,
                            RowCount = span.RowCount,
                            ColumnCount = span.ColumnCount
                        });
                    }
                }
                newSpanDefinitions = spanList.ToArray();
            }

            foreach (var spanDefinition in newSpanDefinitions)
            {
                SheetView.AddColumnHeaderSpanCell(spanDefinition.Row, spanDefinition.Column, spanDefinition.RowCount, spanDefinition.ColumnCount);
            }
        }

        /// <summary>
        /// 列ヘッダー行定義を生成します。
        /// </summary>
        /// <param name="includeProperties">生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>列ヘッダー行定義オブジェクト。</returns>
        public HeaderRowDefinition[] CreateRowDefinitions(string[] includeProperties = null)
        {
            var headerRows = new List<HeaderRowDefinition>();
            foreach (Row row in SheetView.ColumnHeader.Rows)
            {
                var headerRow = new HeaderRowDefinition();

                if (includeProperties == null || includeProperties.Any(x => x.Contains(nameof(Row.Height))))
                {
                    headerRow.Height = row.Height;
                }
                if (includeProperties == null || includeProperties.Any(x => x.Contains(nameof(Row.HorizontalAlignment))))
                {
                    headerRow.HorizontalAlignment = row.HorizontalAlignment;
                }
                if (includeProperties == null || includeProperties.Any(x => x.Contains(nameof(Row.VerticalAlignment))))
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
        /// <param name="includeProperties">生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>列ヘッダーセル定義オブジェクト。</returns>
        public HeaderCellDefinition[][] CreateCellDefinitions(string[] includeProperties = null)
        {
            var allCells = new List<HeaderCellDefinition[]>();
            foreach (Row row in SheetView.ColumnHeader.Rows)
            {
                var cells = new List<HeaderCellDefinition>();
                foreach (Column column in SheetView.ColumnHeader.Columns)
                {
                    var cell = new HeaderCellDefinition();

                    if (includeProperties == null || includeProperties.Any(x => x.Contains(nameof(Cell.Value))))
                    {
                        cell.Value = SheetView.ColumnHeader.Cells[row.Index, column.Index].Value?.ToString();
                    }
                    if (includeProperties == null || includeProperties.Any(x => x.Contains(nameof(Cell.HorizontalAlignment))))
                    {
                        cell.HorizontalAlignment = SheetView.ColumnHeader.Cells[row.Index, column.Index].HorizontalAlignment;
                    }
                    if (includeProperties == null || includeProperties.Any(x => x.Contains(nameof(Cell.VerticalAlignment))))
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
        public SpanDefinition[] CreateSpanDefinitions()
        {
            var headerSpans = new List<SpanDefinition>();
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
                    headerSpans.Add(new SpanDefinition()
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
        /// <param name="templateDefinitionList">テンプレート定義オブジェクト。</param>
        /// <param name="headerRowDefinition">列ヘッダー行定義オブジェクト。</param>
        /// <param name="targetTemplateHeaderRowDefinitionList">対象となったテンプレート列ヘッダー行定義オブジェクト。</param>
        private void ReadRowTemplates(List<TemplateSheetViewDefinition> templateDefinitionList, 
            HeaderRowDefinition headerRowDefinition, List<TemplateHeaderRowDefinition> targetTemplateHeaderRowDefinitionList)
        {
            foreach (var templateDefinition in templateDefinitionList)
            {
                if (templateDefinition.ColumnHeader == null)
                {
                    continue;
                }
                if (templateDefinition.ColumnHeader.Rows == null)
                {
                    continue;
                }

                var templateHeaderRowDefinition = templateDefinition.ColumnHeader.Rows
                    .FirstOrDefault(x => x.TemplateName == headerRowDefinition.BaseTemplate);
                if (templateHeaderRowDefinition == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(templateHeaderRowDefinition.TemplateName))
                {
                    ReadRowTemplates(templateDefinitionList, templateHeaderRowDefinition, targetTemplateHeaderRowDefinitionList);
                }
                targetTemplateHeaderRowDefinitionList.Add(templateHeaderRowDefinition);
            }
        }

        /// <summary>
        /// テンプレートから対象のテンプレート列ヘッダーセル定義オブジェクトを取得する。
        /// </summary>
        /// <param name="templateDefinitionList">テンプレート定義オブジェクト。</param>
        /// <param name="headerCellDefinition">列ヘッダーセル定義オブジェクト。</param>
        /// <param name="targetTemplateHeaderCellDefinitionList">対象となったテンプレート列ヘッダーセル定義オブジェクト。</param>
        private void ReadCellTemplates(List<TemplateSheetViewDefinition> templateDefinitionList,
            HeaderCellDefinition headerCellDefinition, List<TemplateHeaderCellDefinition> targetTemplateHeaderCellDefinitionList)
        {
            foreach (var templateDefinition in templateDefinitionList)
            {
                if (templateDefinition.ColumnHeader == null)
                {
                    continue;
                }
                if (templateDefinition.ColumnHeader.Cells == null)
                {
                    continue;
                }

                var templateHeaderCellDefinition = templateDefinition.ColumnHeader.Cells
                    .FirstOrDefault(x => x.TemplateName == headerCellDefinition.BaseTemplate);
                if (templateHeaderCellDefinition == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(templateHeaderCellDefinition.TemplateName))
                {
                    ReadCellTemplates(templateDefinitionList, templateHeaderCellDefinition, targetTemplateHeaderCellDefinitionList);
                }
                targetTemplateHeaderCellDefinitionList.Add(templateHeaderCellDefinition);
            }
        }

        /// <summary>
        /// 列ヘッダー行のプロパティを設定する。
        /// </summary>
        /// <param name="row">行オブジェクト。</param>
        /// <param name="headerRowDefinition">列ヘッダー行定義オブジェクト。</param>
        private static void SetColumnHeaderRowProps(Row row, HeaderRowDefinition headerRowDefinition)
        {
            if (headerRowDefinition.Height.HasValue)
            {
                row.Height = headerRowDefinition.Height.Value;
            }
            if (headerRowDefinition.HorizontalAlignment.HasValue)
            {
                row.HorizontalAlignment = headerRowDefinition.HorizontalAlignment.Value;
            }
            if (headerRowDefinition.VerticalAlignment.HasValue)
            {
                row.VerticalAlignment = headerRowDefinition.VerticalAlignment.Value;
            }
        }

        /// <summary>
        /// 列ヘッダーセルのプロパティを設定する。
        /// </summary>
        /// <param name="cell">セルオブジェクト。</param>
        /// <param name="headerCellDefinition">列ヘッダーセル定義オブジェクト。</param>
        private static void SetColumnHeaderCellProps(Cell cell, HeaderCellDefinition headerCellDefinition)
        {
            if (!string.IsNullOrEmpty(headerCellDefinition.Value))
            {
                cell.Value = headerCellDefinition.Value;
            }
            if (headerCellDefinition.HorizontalAlignment.HasValue)
            {
                cell.HorizontalAlignment = headerCellDefinition.HorizontalAlignment.Value;
            }
            if (headerCellDefinition.VerticalAlignment.HasValue)
            {
                cell.VerticalAlignment = headerCellDefinition.VerticalAlignment.Value;
            }
        }
    }
}
