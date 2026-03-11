using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Extensions;
using Metroit.Win.GcSpread.Extensions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
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
        /// <param name="columnDefinitions">列定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeJson(ColumnDefinition[] columnDefinitions)
        {
            return JsonConvert.SerializeObject(columnDefinitions);
        }

        /// <summary>
        /// JSON文字列から列定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">列定義JSON文字列。</param>
        /// <returns>列定義オブジェクト。</returns>
        public static ColumnDefinition[] DeserializeJson(string json)
        {
            return JsonConvert.DeserializeObject<ColumnDefinition[]>(json);
        }

        /// <summary>
        /// 列定義オブジェクトから列をセットアップします。
        /// 列は必ず1列目から設定されます。
        /// </summary>
        /// <param name="columnDefinitions">列定義オブジェクト。</param>
        /// <param name="allColumnDefinitions">全列定義オブジェクト。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        public void GenerateColumns(ColumnDefinition[] columnDefinitions, AllColumnDefinition allColumnDefinitions = null,
            Action<Column> generatedColumn = null)
        {
            GenerateColumns(columnDefinitions, null, allColumnDefinitions, generatedColumn);
        }

        /// <summary>
        /// 列定義オブジェクトから列をセットアップします。
        /// 列は必ず1列目から設定されます。
        /// </summary>
        /// <param name="columnDefinitions">列定義オブジェクト。</param>
        /// <param name="templateDefinitionsList">テンプレート定義オブジェクト。</param>
        /// <param name="allColumnDefinitions">全列定義オブジェクト。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        public void GenerateColumns(ColumnDefinition[] columnDefinitions, List<TemplateSheetViewDefinition> templateDefinitionsList,
            AllColumnDefinition allColumnDefinitions, Action<Column> generatedColumn)
        {
            if (columnDefinitions == null)
            {
                return;
            }

            foreach (var column in columnDefinitions.Select((Item, Index) => new { Item, Index }))
            {
                // 列テンプレートの定義を取得する
                var targetTemplates = new List<TemplateColumnDefinition>();
                if (templateDefinitionsList != null)
                {
                    ReadColumnTemplates(templateDefinitionsList, column.Item, targetTemplates);
                }

                // 個別指定のセルタイプとテンプレートのセルタイプに誤差がある場合はエラーとする
                if (targetTemplates
                    .Any(x => string.Compare(x.CellType, column.Item.CellType, true) != 0))
                {
                    throw new ArgumentException("テンプレートとセルタイプが合致していません。", nameof(ColumnDefinition.CellType));
                }

                // 列全体の定義を反映する
                if (allColumnDefinitions != null)
                {
                    SetColumnBaseProps(SheetView.Columns[column.Index], allColumnDefinitions);
                }

                // 列テンプレートから定義を反映する
                foreach (var targetTemplate in targetTemplates)
                {
                    SetColumnProperties(SheetView.Columns[column.Index], targetTemplate);
                }

                // 個別指定された列の定義を反映する
                SetColumnProperties(SheetView.Columns[column.Index], column.Item);

                var cellType = CreateCellType(column.Item.CellType);
                SheetView.Columns[column.Index].CellType = cellType;
                if (cellType != null)
                {
                    // 列テンプレートからセルタイプ定義を反映する
                    foreach (var targetTemplate in targetTemplates)
                    {
                        if (targetTemplate.CellTypeProperties == null)
                        {
                            continue;
                        }
                        cellType.DeserializeJson(JsonConvert.SerializeObject(targetTemplate.CellTypeProperties));
                    }

                    // 個別指定されたセルタイプ定義を反映する
                    if (column.Item.CellTypeProperties != null)
                    {
                        cellType.DeserializeJson(JsonConvert.SerializeObject(column.Item.CellTypeProperties));
                    }
                }

                generatedColumn?.Invoke(SheetView.Columns[column.Index]);
            }
        }

        /// <summary>
        /// 指定したJSONファイルで DataField に合致する列のレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <exception cref="ArgumentException">テンプレートのセルタイプが合致していません。</exception>
        public void GenerateColumnsByDataFieldFrom(string layoutPath, Action<Column> generatedColumn = null)
        {
            GenerateColumnsByDataFieldFrom(layoutPath, string.Empty, generatedColumn);
        }

        /// <summary>
        /// 指定したJSONファイルで DataField に合致する列のレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPath">テンプレートJSONファイルパス。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <exception cref="ArgumentException">テンプレートのセルタイプが合致していません。</exception>
        public void GenerateColumnsByDataFieldFrom(string layoutPath, string templateLayoutPath, Action<Column> generatedColumn = null)
        {
            GenerateColumnsByDataFieldFromInternal(layoutPath, new[] { templateLayoutPath }, generatedColumn);
        }

        /// <summary>
        /// 指定したJSONファイルで DataField に合致する列のレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPaths">テンプレートJSONファイルパス。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <exception cref="ArgumentException">テンプレートのセルタイプが合致していません。</exception>
        private void GenerateColumnsByDataFieldFromInternal(string layoutPath, string[] templateLayoutPaths, Action<Column> generatedColumn)
        {
            string layout = File.ReadAllText(layoutPath);
            var templateLayouts = new List<string>();
            if (templateLayoutPaths.Length > 0)
            {
                foreach (var templateLayoutPath in templateLayoutPaths)
                {
                    if (!string.IsNullOrEmpty(templateLayoutPath))
                    {
                        templateLayouts.Add(File.ReadAllText(templateLayoutPath));
                    }
                }
            }
            GenerateColumnsByDataField(layout, templateLayouts.ToArray(), generatedColumn);
        }

        /// <summary>
        /// DataField に合致する列のレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <exception cref="ArgumentException">テンプレートのセルタイプが合致していません。</exception>
        public void GenerateColumnsByDataField(string layout, Action<Column> generatedColumn = null)
        {
            GenerateColumnsByDataField(layout, string.Empty, generatedColumn);
        }

        /// <summary>
        /// DataField に合致する列のレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="templateLayout">テンプレートJSON文字列。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <exception cref="ArgumentException">テンプレートのセルタイプが合致していません。</exception>
        public void GenerateColumnsByDataField(string layout, string templateLayout, Action<Column> generatedColumn = null)
        {
            GenerateColumnsByDataField(layout, new[] { templateLayout }, generatedColumn);
        }

        /// <summary>
        /// DataField に合致する列のレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="templateLayouts">テンプレートJSON文字列のリスト。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <exception cref="ArgumentException">テンプレートのセルタイプが合致していません。</exception>
        public void GenerateColumnsByDataField(string layout, string[] templateLayouts, Action<Column> generatedColumn = null)
        {
            var columnDefinitions = JsonConvert.DeserializeObject<ColumnDefinition[]>(layout);
            var templateDefinitions = DeserializeTemplateLayouts(templateLayouts);

            GenerateColumnsByDataFieldInternal(columnDefinitions, templateDefinitions, generatedColumn);
        }

        /// <summary>
        /// テンプレートJSONをデシリアライズする。
        /// </summary>
        /// <param name="templateLayouts">テンプレートJSON文字列。</param>
        /// <returns>テンプレートオブジェクトのリスト。1つもテンプレートがなかったときは空のリストを返却する。</returns>
        private static List<TemplateSheetViewDefinition> DeserializeTemplateLayouts(string[] templateLayouts)
        {
            var templateDefinitionsList = new List<TemplateSheetViewDefinition>();
            if (templateLayouts == null)
            {
                return templateDefinitionsList;
            }
            if (templateLayouts.Length == 0)
            {
                return templateDefinitionsList;
            }

            foreach (var templateLayout in templateLayouts)
            {
                if (string.IsNullOrEmpty(templateLayout))
                {
                    continue;
                }

                templateDefinitionsList.Add(JsonConvert.DeserializeObject<TemplateSheetViewDefinition>(templateLayout));
            }

            return templateDefinitionsList;
        }

        /// <summary>
        /// 列定義オブジェクトから列をセットアップします。
        /// 列は必ず1列目から設定されます。
        /// </summary>
        /// <param name="columnDefinitionsArray">列定義オブジェクト。</param>
        /// <param name="templateDefinitionsList">テンプレート定義オブジェクト。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <exception cref="ArgumentException">テンプレートのセルタイプが合致していません。</exception>
        private void GenerateColumnsByDataFieldInternal(ColumnDefinition[] columnDefinitionsArray,
            List<TemplateSheetViewDefinition> templateDefinitionsList, Action<Column> generatedColumn)
        {
            if (columnDefinitionsArray == null)
            {
                return;
            }

            foreach (var columnDefinitions in columnDefinitionsArray)
            {
                var columnIndex = SheetView.GetColumnIndexFromDataField(columnDefinitions.DataField);

                // 列テンプレートの定義を取得する
                var targetTemplatesList = new List<TemplateColumnDefinition>();
                if (templateDefinitionsList != null)
                {
                    ReadColumnTemplates(templateDefinitionsList, columnDefinitions, targetTemplatesList);
                }

                // 個別指定のセルタイプとテンプレートのセルタイプに誤差がある場合はエラーとする
                if (targetTemplatesList
                    .Any(x => string.Compare(x.CellType, columnDefinitions.CellType, true) != 0))
                {
                    throw new ArgumentException("テンプレートとセルタイプが合致していません。", nameof(ColumnDefinition.CellType));
                }

                // 列テンプレートから定義を反映する
                foreach (var targetTemplate in targetTemplatesList)
                {
                    SetColumnProperties(SheetView.Columns[columnIndex], targetTemplate);
                }

                // 個別指定された列の定義を反映する
                SetColumnProperties(SheetView.Columns[columnIndex], columnDefinitions);

                ICellType cellType;
                if (SheetView.Columns[columnIndex].CellType.GetType().Name == columnDefinitions.CellType)
                {
                    cellType = SheetView.Columns[columnIndex].CellType;
                }
                else
                {
                    cellType = CreateCellType(columnDefinitions.CellType);
                    SheetView.Columns[columnIndex].CellType = cellType;
                }

                if (cellType != null)
                {
                    // 列テンプレートからセルタイプ定義を反映する
                    foreach (var targetTemplate in targetTemplatesList)
                    {
                        if (targetTemplate.CellTypeProperties == null)
                        {
                            continue;
                        }
                        cellType.DeserializeJson(JsonConvert.SerializeObject(targetTemplate.CellTypeProperties));
                    }

                    // 個別指定されたセルタイプ定義を反映する
                    if (columnDefinitions.CellTypeProperties != null)
                    {
                        cellType.DeserializeJson(JsonConvert.SerializeObject(columnDefinitions.CellTypeProperties));
                    }
                }

                generatedColumn?.Invoke(SheetView.Columns[columnIndex]);
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
        public ColumnDefinition[] CreateColumnsDefinitions(Func<Column, Dictionary<string, object>> columnOptions,
            string[] includeProps = null, bool outputCellTypeProps = true, bool outputOptions = true)
        {
            var columns = new List<ColumnDefinition>();

            foreach (Column svColumn in SheetView.Columns)
            {
                var column = new ColumnDefinition();

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
                    column.CellTypeProperties = JObject.Parse(svColumn.CellType.SerializeJson());
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
        /// <param name="columnDefinitions">列定義オブジェクト。</param>
        private void SetColumnBaseProps(Column column, ColumnDefinitionBase columnDefinitions)
        {
            if (columnDefinitions == null)
            {
                return;
            }

            if (columnDefinitions.AllowAutoFilter.HasValue)
            {
                column.AllowAutoFilter = columnDefinitions.AllowAutoFilter.Value;
            }
            if (columnDefinitions.AllowAutoSort.HasValue)
            {
                column.AllowAutoSort = columnDefinitions.AllowAutoSort.Value;
            }
            if (columnDefinitions.HorizontalAlignment.HasValue)
            {
                column.HorizontalAlignment = columnDefinitions.HorizontalAlignment.Value;
            }
            if (columnDefinitions.ImeMode.HasValue)
            {
                column.ImeMode = columnDefinitions.ImeMode.Value;
            }
            if (columnDefinitions.Locked.HasValue)
            {
                column.Locked = columnDefinitions.Locked.Value;
            }
            if (columnDefinitions.VerticalAlignment.HasValue)
            {
                column.VerticalAlignment = columnDefinitions.VerticalAlignment.Value;
            }
            if (columnDefinitions.Visible.HasValue)
            {
                column.Visible = columnDefinitions.Visible.Value;
            }
            if (columnDefinitions.Width.HasValue)
            {
                column.Width = columnDefinitions.Width.Value;
            }
        }

        /// <summary>
        /// 列の定義を反映する。
        /// </summary>
        /// <param name="column">列オブジェクト。</param>
        /// <param name="columnDefinitions">列定義オブジェクト。</param>
        private void SetColumnProperties(Column column, ColumnDefinition columnDefinitions)
        {
            if (columnDefinitions == null)
            {
                return;
            }

            SetColumnBaseProps(column, columnDefinitions);

            if (!string.IsNullOrEmpty(columnDefinitions.DataField))
            {
                SheetView.BindDataColumn(column.Index, columnDefinitions.DataField);
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
            if (string.Compare(cellTypeName, nameof(RegularExpressionCellType), true) == 0)
            {
                return new RegularExpressionCellType();
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
            if (string.Compare(cellTypeName, nameof(GcNumberCellType), true) == 0)
            {
                return new GcNumberCellType();
            }
            if (string.Compare(cellTypeName, nameof(GcComboBoxCellType), true) == 0)
            {
                return new GcComboBoxCellType();
            }

            return null;
        }

        /// <summary>
        /// 全テンプレート情報から対象の列定義テンプレート情報を取得する。
        /// </summary>
        /// <param name="templateDefinitionsList">テンプレート定義オブジェクト。</param>
        /// <param name="columnDefinitions">列定義オブジェクト。</param>
        /// <param name="targetTemplateColumnDefinitionsList">対象となった列定義オブジェクト。</param>
        private static void ReadColumnTemplates(List<TemplateSheetViewDefinition> templateDefinitionsList,
            ColumnDefinition columnDefinitions, List<TemplateColumnDefinition> targetTemplateColumnDefinitionsList)
        {
            foreach (var templateDefinitions in templateDefinitionsList)
            {
                if (templateDefinitions.Columns == null)
                {
                    continue;
                }

                var templateColumnDefinitions = templateDefinitions.Columns
                    .FirstOrDefault(x => x.TemplateName == columnDefinitions.BaseTemplate);
                if (templateColumnDefinitions == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(templateColumnDefinitions.TemplateName))
                {
                    ReadColumnTemplates(templateDefinitionsList, templateColumnDefinitions, targetTemplateColumnDefinitionsList);
                }
                targetTemplateColumnDefinitionsList.Add(templateColumnDefinitions);
            }
        }
    }
}
