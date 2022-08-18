using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// Column のセットアップ操作を提供します。
    /// </summary>
    public class ColumnsSetupProvider
    {
        /// <summary>
        /// 処理の対象となる SheetView。
        /// </summary>
        private SheetView SheetView { get; }

        /// <summary>
        /// 新しい ColumnSetupProvider インスタンスを生成します。
        /// </summary>
        /// <param name="sheetView">処理対象とする SheetView。</param>
        public ColumnsSetupProvider(SheetView sheetView)
        {
            SheetView = sheetView;
        }

        /// <summary>
        /// 列定義から列をセットアップします。
        /// 列は必ず1列目から設定されます。
        /// </summary>
        /// <param name="columnDefs">列定義オブジェクト。</param>
        /// <param name="allColumnDefs">全列定義オブジェクト。</param>
        /// <param name="templateLayoutDefsList">テンプレート定義オブジェクト。</param>
        public void SetupColumns(ColumnDefinitions[] columnDefs, AllColumnDefinitions allColumnDefs = null, List<TemplateSheetViewDefinitions> templateLayoutDefsList = null)
        {
            if (columnDefs == null)
            {
                return;
            }

            foreach (var column in columnDefs.Select((Item, Index) => new { Item, Index }))
            {
                // 列テンプレートの定義を取得する
                var targetTemplates = new List<TemplateColumnDefinitions>();
                if (templateLayoutDefsList != null)
                {
                    ReadColumnTemplates(templateLayoutDefsList, column.Item, targetTemplates);
                }

                // 個別指定のセルタイプとテンプレートのセルタイプに誤差がある場合はエラーとする
                if (targetTemplates.Where(x => string.Compare(x.CellType, column.Item.CellType, true) != 0).Count() > 0)
                {
                    throw new ArgumentException("テンプレートとセルタイプが合致していません。", nameof(ColumnDefinitions.CellType));
                }

                // 列全体の定義を反映する
                if (allColumnDefs != null)
                {
                    SetColumnBaseProps(SheetView.Columns[column.Index], allColumnDefs);
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
                        cellType.DeserializeJson(Newtonsoft.Json.JsonConvert.SerializeObject(targetTemplate.CellTypeProps));
                    }

                    // 個別指定されたセルタイプ定義を反映する
                    if (column.Item.CellTypeProps != null)
                    {
                        cellType.DeserializeJson(Newtonsoft.Json.JsonConvert.SerializeObject(column.Item.CellTypeProps));
                    }
                }
            }
        }

        /// <summary>
        /// 列全体の定義を反映する。
        /// </summary>
        /// <param name="column">列オブジェクト。</param>
        /// <param name="allColumnDefs">列定義オブジェクト。</param>
        private void SetColumnBaseProps(Column column, ColumnDefinitionsBase allColumnDefs)
        {
            if (allColumnDefs == null)
            {
                return;
            }

            if (allColumnDefs.AllowAutoFilter.HasValue)
            {
                column.AllowAutoFilter = allColumnDefs.AllowAutoFilter.Value;
            }
            if (allColumnDefs.AllowAutoSort.HasValue)
            {
                column.AllowAutoSort = allColumnDefs.AllowAutoSort.Value;
            }
            if (allColumnDefs.HorizontalAlignment.HasValue)
            {
                column.HorizontalAlignment = allColumnDefs.HorizontalAlignment.Value;
            }
            if (allColumnDefs.ImeMode.HasValue)
            {
                column.ImeMode = allColumnDefs.ImeMode.Value;
            }
            if (allColumnDefs.Locked.HasValue)
            {
                column.Locked = allColumnDefs.Locked.Value;
            }
            if (allColumnDefs.VerticalAlignment.HasValue)
            {
                column.VerticalAlignment = allColumnDefs.VerticalAlignment.Value;
            }
            if (allColumnDefs.Visible.HasValue)
            {
                column.Visible = allColumnDefs.Visible.Value;
            }
            if (allColumnDefs.Width.HasValue)
            {
                column.Width = allColumnDefs.Width.Value;
            }
        }

        /// <summary>
        /// 列の定義を反映する。
        /// </summary>
        /// <param name="column">列オブジェクト。</param>
        /// <param name="columnDefs">列定義オブジェクト。</param>
        private void SetColumnProps(Column column, ColumnDefinitions columnDefs)
        {
            if (columnDefs == null)
            {
                return;
            }

            SetColumnBaseProps(column, columnDefs);

            if (!string.IsNullOrEmpty(columnDefs.DataField))
            {
                SheetView.BindDataColumn(column.Index, columnDefs.DataField);
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
        /// 列定義を生成します。
        /// </summary>
        /// <param name="columnOptions">列のオプション情報設定アクションデリゲート。</param>
        /// <returns>ColumnDefinitions[] オブジェクト。</returns>
        public ColumnDefinitions[] CreateColumnsDefinitions(Func<Column, Dictionary<string, object>> columnOptions)
        {
            var columns = new List<ColumnDefinitions>();

            foreach (Column svColumn in SheetView.Columns)
            {
                //// TODO: GcComboBoxのみ動くようにする
                //if (!(svColumn.CellType is GcComboBoxCellType) && !(svColumn.CellType is GcNumberCellType) && !(svColumn.CellType is GcTextBoxCellType) && !(svColumn.CellType is GcDateTimeCellType))
                //{
                //    continue;
                //}


                var column = new ColumnDefinitions()
                {
                    DataField = svColumn.DataField,
                    HorizontalAlignment = svColumn.HorizontalAlignment,
                    VerticalAlignment = svColumn.VerticalAlignment,
                    AllowAutoFilter = svColumn.AllowAutoFilter,
                    AllowAutoSort = svColumn.AllowAutoSort,
                    Width = svColumn.Width,
                    Visible = svColumn.Visible,
                    CellType = svColumn.CellType.ToString(),
                    ImeMode = svColumn.ImeMode,
                    Locked = svColumn.Locked
                };

                //if (svColumn.CellType is TextCellType)
                //{
                //    column.CellTypeProps = DecompileTextCellTypeProperties((TextCellType)svColumn.CellType);
                //}
                //if (svColumn.CellType is DateTimeCellType)
                //{
                //    column.CellTypeProps = DecompileDateTimeCellTypeProperties((DateTimeCellType)svColumn.CellType);
                //}
                //if (svColumn.CellType is NumberCellType)
                //{
                //    column.CellTypeProps = DecompileNumberCellTypeProperties((NumberCellType)svColumn.CellType);
                //}
                //if (svColumn.CellType is ComboBoxCellType)
                //{
                //    column.CellTypeProps = DecompileComboBoxCellTypeProperties((ComboBoxCellType)svColumn.CellType);
                //}
                //if (svColumn.CellType is CheckBoxCellType)
                //{
                //    column.CellTypeProps = DecompileCheckBoxCellTypeProperties((CheckBoxCellType)svColumn.CellType);
                //}
                //if (svColumn.CellType is ButtonCellType)
                //{
                //    column.CellTypeProps = DecompileButtonCellTypeProperties((ButtonCellType)svColumn.CellType);
                //}
                //if (svColumn.CellType is GcTextBoxCellType)
                //{
                //    column.CellTypeProps = DecompileGcTextBoxCellTypeProperties((GcTextBoxCellType)svColumn.CellType);
                //}
                //if (svColumn.CellType is GcDateTimeCellType)
                //{
                //    column.CellTypeProps = DecompileGcDateTimeCellTypeProperties((GcDateTimeCellType)svColumn.CellType);
                //}

                column.CellTypeProps = svColumn.CellType.SerializeJson();

                column.Options = columnOptions?.Invoke(svColumn);

                columns.Add(column);
            }

            return columns.ToArray();
        }

        /// <summary>
        /// 全テンプレート情報から対象の列定義テンプレート情報を取得する。
        /// </summary>
        /// <param name="templateLayoutDefsList">テンプレート定義オブジェクト。</param>
        /// <param name="columnDefinitions">列定義オブジェクト。</param>
        /// <param name="targetTemplates">対象となった列定義オブジェクト。</param>
        private static void ReadColumnTemplates(List<TemplateSheetViewDefinitions> templateLayoutDefsList, ColumnDefinitions columnDefinitions, List<TemplateColumnDefinitions> targetTemplates)
        {
            foreach (var templateLayoutDefs in templateLayoutDefsList)
            {
                var templateColumnDefs = templateLayoutDefs.Columns.Where(x => x.TemplateName == columnDefinitions.BaseTemplate).FirstOrDefault();
                if (templateColumnDefs == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(templateColumnDefs.TemplateName))
                {
                    ReadColumnTemplates(templateLayoutDefsList, templateColumnDefs, targetTemplates);
                }
                targetTemplates.Add(templateColumnDefs);
            }
        }
    }
}
