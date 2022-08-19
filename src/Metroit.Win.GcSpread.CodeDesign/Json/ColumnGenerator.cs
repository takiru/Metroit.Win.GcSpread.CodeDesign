using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Extensions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// Column の生成操作を提供します。
    /// </summary>
    public class ColumnGenerator
    {
        /// <summary>
        /// 処理の対象となる SheetView。
        /// </summary>
        private SheetView SheetView { get; }

        /// <summary>
        /// 新しい ColumnGenerator インスタンスを生成します。
        /// </summary>
        /// <param name="sheetView">処理対象とする SheetView。</param>
        public ColumnGenerator(SheetView sheetView)
        {
            SheetView = sheetView;
        }

        /// <summary>
        /// 列定義オブジェクトからJSON文字列にシリアライズします。
        /// </summary>
        /// <param name="defs">列定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeJson(ColumnDefinitions[] defs)
        {
            return JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// JSON文字列から列定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列定義JSON文字列。</param>
        /// <returns>列定義オブジェクト。</returns>
        public static ColumnDefinitions[] DeserializeJson(string json)
        {
            return JsonConvert.DeserializeObject<ColumnDefinitions[]>(json);
        }

        /// <summary>
        /// 列定義オブジェクトから列をセットアップします。
        /// 列は必ず1列目から設定されます。
        /// </summary>
        /// <param name="defs">列定義オブジェクト。</param>
        /// <param name="allColumnDef">全列定義オブジェクト。</param>
        public void GenerateColumns(ColumnDefinitions[] defs, AllColumnDefinitions allColumnDef = null)
        {
            GenerateColumns(defs, null, allColumnDef);
        }

        /// <summary>
        /// 列定義オブジェクトから列をセットアップします。
        /// 列は必ず1列目から設定されます。
        /// </summary>
        /// <param name="defs">列定義オブジェクト。</param>
        /// <param name="templateDefs">テンプレート定義オブジェクト。</param>
        /// <param name="allColumnDef">全列定義オブジェクト。</param>
        public void GenerateColumns(ColumnDefinitions[] defs, List<TemplateSheetViewDefinitions> templateDefs, AllColumnDefinitions allColumnDef = null)
        {
            if (defs == null)
            {
                return;
            }

            foreach (var column in defs.Select((Item, Index) => new { Item, Index }))
            {
                // 列テンプレートの定義を取得する
                var targetTemplates = new List<TemplateColumnDefinitions>();
                if (templateDefs != null)
                {
                    ReadColumnTemplates(templateDefs, column.Item, targetTemplates);
                }

                // 個別指定のセルタイプとテンプレートのセルタイプに誤差がある場合はエラーとする
                if (targetTemplates.Where(x => string.Compare(x.CellType, column.Item.CellType, true) != 0).Count() > 0)
                {
                    throw new ArgumentException("テンプレートとセルタイプが合致していません。", nameof(ColumnDefinitions.CellType));
                }

                // 列全体の定義を反映する
                if (allColumnDef != null)
                {
                    SetColumnBaseProps(SheetView.Columns[column.Index], allColumnDef);
                }

                // 列テンプレートから定義を反映する
                foreach (var targetTemplate in targetTemplates)
                {
                    SetColumnProps(SheetView.Columns[column.Index], targetTemplate);
                }

                // 個別指定された列の定義を反映する
                SetColumnProps(SheetView.Columns[column.Index], column.Item);

                ICellType cellType = CreateCellType(column.Item.CellType);
                SheetView.Columns[column.Index].CellType = cellType;
                if (cellType != null)
                {
                    // 列テンプレートからセルタイプ定義を反映する
                    foreach (var targetTemplate in targetTemplates)
                    {
                        if (targetTemplate.CellTypeProps == null)
                        {
                            continue;
                        }
                        cellType.DeserializeJson(JsonConvert.SerializeObject(targetTemplate.CellTypeProps));
                    }

                    // 個別指定されたセルタイプ定義を反映する
                    if (column.Item.CellTypeProps != null)
                    {
                        cellType.DeserializeJson(JsonConvert.SerializeObject(column.Item.CellTypeProps));
                    }
                }
            }
        }

        /// <summary>
        /// 列定義オブジェクトを生成します。
        /// </summary>
        /// <param name="columnOptions">列のオプション情報設定アクションデリゲート。</param>
        /// <param name="includeProps">生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="outputCellTypeProps">セルタイププロパティを生成するかどうか。</param>
        /// <param name="outputOptions">オプション情報を生成するかどうか。false の場合、第一引数は動作しません。</param>
        /// <returns>ColumnDefinitions[] オブジェクト。</returns>
        public ColumnDefinitions[] CreateColumnsDefinitions(Func<Column, Dictionary<string, object>> columnOptions,
            string[] includeProps = null, bool outputCellTypeProps = true, bool outputOptions = true)
        {
            var columns = new List<ColumnDefinitions>();

            foreach (Column svColumn in SheetView.Columns)
            {
                var column = new ColumnDefinitions();

                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.DataField))))
                {
                    column.DataField = svColumn.DataField;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.HorizontalAlignment))))
                {
                    column.HorizontalAlignment = svColumn.HorizontalAlignment;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.VerticalAlignment))))
                {
                    column.VerticalAlignment = svColumn.VerticalAlignment;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.AllowAutoFilter))))
                {
                    column.AllowAutoFilter = svColumn.AllowAutoFilter;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.AllowAutoSort))))
                {
                    column.AllowAutoSort = svColumn.AllowAutoSort;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.Width))))
                {
                    column.Width = svColumn.Width;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.Visible))))
                {
                    column.Visible = svColumn.Visible;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.CellType))))
                {
                    column.CellType = svColumn.CellType.ToString();
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.ImeMode))))
                {
                    column.ImeMode = svColumn.ImeMode;
                }
                if (includeProps == null || includeProps.Any(x => x.Contains(nameof(Column.Locked))))
                {
                    column.Locked = svColumn.Locked;
                }

                if (outputCellTypeProps)
                {
                    column.CellTypeProps = JObject.Parse(svColumn.CellType.SerializeJson());
                }
                if (outputOptions)
                {
                    column.Options = columnOptions?.Invoke(svColumn);
                }

                columns.Add(column);
            }

            return columns.ToArray();
        }

        /// <summary>
        /// 列全体の定義を反映する。
        /// </summary>
        /// <param name="column">列オブジェクト。</param>
        /// <param name="def">列定義オブジェクト。</param>
        private void SetColumnBaseProps(Column column, ColumnDefinitionsBase def)
        {
            if (def == null)
            {
                return;
            }

            if (def.AllowAutoFilter.HasValue)
            {
                column.AllowAutoFilter = def.AllowAutoFilter.Value;
            }
            if (def.AllowAutoSort.HasValue)
            {
                column.AllowAutoSort = def.AllowAutoSort.Value;
            }
            if (def.HorizontalAlignment.HasValue)
            {
                column.HorizontalAlignment = def.HorizontalAlignment.Value;
            }
            if (def.ImeMode.HasValue)
            {
                column.ImeMode = def.ImeMode.Value;
            }
            if (def.Locked.HasValue)
            {
                column.Locked = def.Locked.Value;
            }
            if (def.VerticalAlignment.HasValue)
            {
                column.VerticalAlignment = def.VerticalAlignment.Value;
            }
            if (def.Visible.HasValue)
            {
                column.Visible = def.Visible.Value;
            }
            if (def.Width.HasValue)
            {
                column.Width = def.Width.Value;
            }
        }

        /// <summary>
        /// 列の定義を反映する。
        /// </summary>
        /// <param name="column">列オブジェクト。</param>
        /// <param name="def">列定義オブジェクト。</param>
        private void SetColumnProps(Column column, ColumnDefinitions def)
        {
            if (def == null)
            {
                return;
            }

            SetColumnBaseProps(column, def);

            if (!string.IsNullOrEmpty(def.DataField))
            {
                SheetView.BindDataColumn(column.Index, def.DataField);
            }
        }

        /// <summary>
        /// セルタイプを生成する。
        /// </summary>
        /// <param name="cellTypeName">セルタイプ名。</param>
        /// <returns>セルタイプ。</returns>
        private ICellType CreateCellType(string cellTypeName)
        {
            // 通常セルタイプ
            if (string.Compare(cellTypeName, nameof(TextCellType), true) == 0)
            {
                return new TextCellType();
            }
            if (string.Compare(cellTypeName, nameof(DateTimeCellType), true) == 0)
            {
                return new DateTimeCellType();
            }
            if (string.Compare(cellTypeName, nameof(NumberCellType), true) == 0)
            {
                return new NumberCellType();
            }
            if (string.Compare(cellTypeName, nameof(ComboBoxCellType), true) == 0)
            {
                return new ComboBoxCellType();
            }
            if (string.Compare(cellTypeName, nameof(CheckBoxCellType), true) == 0)
            {
                return new CheckBoxCellType();
            }
            if (string.Compare(cellTypeName, nameof(ButtonCellType), true) == 0)
            {
                return new ButtonCellType();
            }

            // InputManセルタイプ
            if (string.Compare(cellTypeName, nameof(GcTextBoxCellType), true) == 0)
            {
                return new GcTextBoxCellType();
            }
            if (string.Compare(cellTypeName, nameof(GcDateTimeCellType), true) == 0)
            {
                return new GcDateTimeCellType();
            }

            return null;
        }

        /// <summary>
        /// 全テンプレート情報から対象の列定義テンプレート情報を取得する。
        /// </summary>
        /// <param name="templateDefs">テンプレート定義オブジェクト。</param>
        /// <param name="def">列定義オブジェクト。</param>
        /// <param name="targetTemplateDefs">対象となった列定義オブジェクト。</param>
        private static void ReadColumnTemplates(List<TemplateSheetViewDefinitions> templateDefs, ColumnDefinitions def, List<TemplateColumnDefinitions> targetTemplateDefs)
        {
            foreach (var templateDef in templateDefs)
            {
                var columnDef = templateDef.Columns.Where(x => x.TemplateName == def.BaseTemplate).FirstOrDefault();
                if (columnDef == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(columnDef.TemplateName))
                {
                    ReadColumnTemplates(templateDefs, columnDef, targetTemplateDefs);
                }
                targetTemplateDefs.Add(columnDef);
            }
        }
    }
}
