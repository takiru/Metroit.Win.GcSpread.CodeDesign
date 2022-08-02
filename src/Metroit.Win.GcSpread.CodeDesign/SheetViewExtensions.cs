﻿using Metroit.Win.GcSpread.CodeDesign.Json;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using GrapeCity.Win.Spread.InputMan.CellType;
using LengthUnit = GrapeCity.Win.Spread.InputMan.CellType.LengthUnit;
using System.Drawing;
using GrapeCity.Win.Spread.InputMan.CellType.Fields;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// GrapeCity SPREAD SheetView の拡張メソッドを提供します。
    /// </summary>
    public static class SheetViewExtensions
    {
        /// <summary>
        /// 指定したJSONファイルでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="path">JSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayout(this SheetView sheetView, string path,
            string templateColumnsPath = "",
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            var jsonObj = JsonConvert.DeserializeSheetView(File.ReadAllText(path));
            TemplateColumnsDefinitions templateColumnsObj = null;
            if (!string.IsNullOrEmpty(templateColumnsPath))
            {
                templateColumnsObj = Newtonsoft.Json.JsonConvert.DeserializeObject<TemplateColumnsDefinitions>(File.ReadAllText(templateColumnsPath));
            }
            BindJsonLayout(sheetView, jsonObj, templateColumnsObj, columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONオブジェクトでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="jsonObj">バインドするJSONオブジェクト。</param>
        /// <param name="templateColumnsObj">列テンプレートJSONオブジェクト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayout(this SheetView sheetView, SheetViewDefinitions jsonObj,
            TemplateColumnsDefinitions templateColumnsObj = null,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            // バインド準備
            PrepareCompilation(sheetView);

            // ヘッダーを編集
            ComplieColumnHeader(sheetView, jsonObj.ColumnHeader, columnHeaderCellTag);

            // 列を編集
            CompileColumns(sheetView, jsonObj, templateColumnsObj);

            // JSONデータを超えて列移動やらタイトル変更やらを行いたい場合のインターセプター
            var bindResults = interceptor?.Invoke(sheetView, jsonObj);

            // ヘッダースパンを編集
            CompileColumnHeaderSpans(sheetView, jsonObj.ColumnHeader.Spans, bindResults);
        }

        /// <summary>
        /// 指定したJSONファイルでシートのヘッダー部のみをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="path">JSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayoutWithoutColumns(this SheetView sheetView, string path,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            var jsonObj = JsonConvert.DeserializeSheetView(File.ReadAllText(path));
            BindJsonLayoutWithoutColumns(sheetView, jsonObj, columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONオブジェクトでシートのヘッダー部のみをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="jsonObj">バインドするJSONオブジェクト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayoutWithoutColumns(this SheetView sheetView, SheetViewDefinitions jsonObj,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            // バインド準備
            PrepareCompilation(sheetView);

            // ヘッダーを編集
            ComplieColumnHeader(sheetView, jsonObj.ColumnHeader, columnHeaderCellTag);

            // JSONデータを超えて列移動やらタイトル変更やらを行いたい場合のインターセプター
            var bindResults = interceptor?.Invoke(sheetView, jsonObj);

            // ヘッダースパンを編集
            CompileColumnHeaderSpans(sheetView, jsonObj.ColumnHeader.Spans, bindResults);
        }

        /// <summary>
        /// DataField プロパティを保持したまま列の移動を行います。
        /// 列タイトルのセル結合を保持するため、SheetView.MoveColumn() の第四引数 moveContent は true として実行されます。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="fromIndex">移動元の列インデックス。</param>
        /// <param name="toIndex">移動先の列インデックス。</param>
        /// <param name="columnCount">移動する列数。</param>
        /// <returns>列の移動を行った情報のリスト。移動が行われなかった時は空のリストを返却します。</returns>
        public static List<ColumnMoveResult> MoveColumnKeepDataField(this SheetView sheetView,
            int fromIndex, int toIndex, int columnCount)
        {
            var result = new List<ColumnMoveResult>();

            // 移動する必要がない
            if (fromIndex == toIndex)
            {
                return result;
            }

            // 列を後ろに移動する時
            if (fromIndex < toIndex)
            {
                // 移動を指定した列分を把握
                // NOTE: MoveColumn() の仕様により、columnCount > 1 の時、toIndex の位置は、移動する列の最後の列となる
                var lastToIndex = 0;
                for (var i = fromIndex; i < fromIndex + columnCount; i++)
                {
                    lastToIndex = (toIndex - columnCount) + (i - fromIndex) + 1;
                    result.Add(new ColumnMoveResult(i, i, lastToIndex, sheetView.Columns[i].DataField));
                }

                // 移動することによって前に移動する列分を把握
                var intervalStartInex = fromIndex + columnCount;
                for (var i = intervalStartInex; i <= lastToIndex; i++)
                {
                    result.Add(new ColumnMoveResult(i, i, i - columnCount, sheetView.Columns[i].DataField));
                }
            }

            // 列を前に移動する時
            if (fromIndex > toIndex)
            {
                // 移動を指定した列分を把握
                // NOTE: MoveColumn() の仕様により、columnCount > 1 の時、toIndex の位置は、移動する列の最初の列となる
                for (var i = fromIndex; i < fromIndex + columnCount; i++)
                {
                    var lastToIndex = toIndex + (i - fromIndex);
                    result.Add(new ColumnMoveResult(i, i, lastToIndex, sheetView.Columns[i].DataField));
                }

                // 移動することによって後ろに移動する列分を把握
                var intervalEndIndex = fromIndex - 1;
                for (var i = toIndex; i <= intervalEndIndex; i++)
                {
                    result.Add(new ColumnMoveResult(i, i, i + columnCount, sheetView.Columns[i].DataField));
                }
            }

            // 実際の列移動とバインドし直し
            sheetView.MoveColumn(fromIndex, toIndex, columnCount, true);
            foreach (var columnMoveResult in result)
            {
                sheetView.BindDataColumn(columnMoveResult.AfterColumnIndex, columnMoveResult.DataField);
            }

            return result;
        }

        /// <summary>
        /// シート情報からJSONデータオブジェクトを生成します。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="columnOptions">列のオプション情報設定アクションデリゲート。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public static SheetViewDefinitions CreateSheetDefinitions(this SheetView sheetView, Func<Column, Dictionary<string, object>> columnOptions = null)
        {
            var result = new SheetViewDefinitions()
            {
                ColumnHeader = new Json.ColumnHeaderDefinitions()
            };

            // ヘッダー行/列情報
            result.ColumnHeader.Rows = DecompileColumnHeaderRows(sheetView);

            // ヘッダースパン情報
            result.ColumnHeader.Spans = DecompileColumnHeaderSpans(sheetView);

            // 列情報
            result.Columns = DecompileColumns(sheetView, columnOptions);

            return result;
        }

        /// <summary>
        /// シート情報から列ヘッダーの特定のみのJSONデータオブジェクトを生成します。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public static SpecialColumnHeaderDefinitions CreateSpecialColumnHeaderDefinitions(this SheetView sheetView)
        {
            var result = new SpecialColumnHeaderDefinitions();

            var headerRows = new List<SpecialHeaderRowDefinitions>();
            foreach (Row row in sheetView.ColumnHeader.Rows)
            {
                var headerRow = new SpecialHeaderRowDefinitions();

                var headerColumns = new List<SpecialHeaderColumnDefinitions>();
                foreach (Column column in sheetView.ColumnHeader.Columns)
                {
                    headerColumns.Add(new SpecialHeaderColumnDefinitions()
                    {
                        Value = sheetView.ColumnHeader.Cells[row.Index, column.Index].Value?.ToString()
                    });
                }
                headerRow.Columns = headerColumns.ToArray();
                headerRows.Add(headerRow);
            }
            result.Rows = headerRows.ToArray();

            return result;
        }

        /// <summary>
        /// シート情報から列の特定のみのJSONデータオブジェクトを生成します。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public static SpecialColumnsDefinitions CreateSpecialColumnDefinitions(this SheetView sheetView)
        {
            var result = new SpecialColumnsDefinitions();

            var columns = new List<SpecialColumnDefinitions>();

            foreach (Column svColumn in sheetView.Columns)
            {
                var column = new SpecialColumnDefinitions()
                {
                    Width = svColumn.Width,
                    Visible = svColumn.Visible,
                    DataField = svColumn.DataField
                };
                columns.Add(column);
            }
            result.Columns = columns.ToArray();

            return result;
        }

        /// <summary>
        /// バインド準備を行う。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        private static void PrepareCompilation(SheetView sheetView)
        {
            // NOTE: レイアウト読込において無効にしなければならないプロパティが有効な場合は実行不可とする
            if (sheetView.AutoGenerateColumns)
            {
                throw new InvalidOperationException("AutoGenerateColumns = true の場合、バインドを行うことはできません。");
            }
            if (sheetView.DataAutoCellTypes)
            {
                throw new InvalidOperationException("DataAutoCellTypes = true の場合、バインドを行うことはできません。");
            }
            if (sheetView.DataAutoSizeColumns)
            {
                throw new InvalidOperationException("DataAutoSizeColumns = true の場合、バインドを行うことはできません。");
            }
            if (sheetView.DataAutoHeadings)
            {
                throw new InvalidOperationException("DataAutoHeadings = true の場合、バインドを行うことはできません。");
            }

            sheetView.ColumnHeader.Columns.Count = 0;
            sheetView.ColumnHeader.Rows.Count = 0;
            sheetView.Columns.Count = 0;
        }

        /// <summary>
        /// ヘッダー情報をコンパイルする。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="columnHeader">Json.ColumnHeader オブジェクト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        private static void ComplieColumnHeader(SheetView sheetView, ColumnHeaderDefinitions columnHeader, Func<Cell, object> columnHeaderCellTag)
        {
            const string headerTagPrefix = @"Header";

            // ヘッダーの列・行を生成
            sheetView.ColumnHeader.Columns.Add(0, columnHeader.Rows.Max(x => x.Columns.Length));
            sheetView.ColumnHeader.Rows.Add(0, columnHeader.Rows.Length);
            foreach (var row in columnHeader.Rows.Select((Item, Index) => new { Item, Index }))
            {
                if (row.Item.Height.HasValue)
                {
                    sheetView.ColumnHeader.Rows[row.Index].Height = row.Item.Height.Value;
                }
                if (row.Item.HorizontalAlignment.HasValue)
                {
                    sheetView.ColumnHeader.Rows[row.Index].HorizontalAlignment = row.Item.HorizontalAlignment.Value;
                }
                if (row.Item.VerticalAlignment.HasValue)
                {
                    sheetView.ColumnHeader.Rows[row.Index].VerticalAlignment = row.Item.VerticalAlignment.Value;
                }

                foreach (var column in row.Item.Columns.Select((Item, Index) => new { Item, Index }))
                {
                    if (columnHeaderCellTag == null)
                    {
                        sheetView.ColumnHeader.Cells[row.Index, column.Index].Tag = $"{headerTagPrefix}_{row.Index}_{column.Index}";
                    }
                    else
                    {
                        sheetView.ColumnHeader.Cells[row.Index, column.Index].Tag = columnHeaderCellTag.Invoke(sheetView.ColumnHeader.Cells[row.Index, column.Index]);
                    }

                    if (!string.IsNullOrEmpty(column.Item.Value))
                    {
                        sheetView.ColumnHeader.Cells[row.Index, column.Index].Value = column.Item.Value;
                    }
                    if (column.Item.HorizontalAlignment.HasValue)
                    {
                        sheetView.ColumnHeader.Cells[row.Index, column.Index].HorizontalAlignment = column.Item.HorizontalAlignment.Value;
                    }
                    if (column.Item.VerticalAlignment.HasValue)
                    {
                        sheetView.ColumnHeader.Cells[row.Index, column.Index].VerticalAlignment = column.Item.VerticalAlignment.Value;
                    }
                }
            }
        }

        /// <summary>
        /// ヘッダースパン情報をコンパイルする。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="spans">Span[] オブジェクト。</param>
        /// <param name="columnMoveResults">列の移動を行った情報のリスト。</param>
        private static void CompileColumnHeaderSpans(SheetView sheetView, SpanDefinitions[] spans, IList<ColumnMoveResult> columnMoveResults)
        {
            SpanDefinitions[] newSpans;

            if (columnMoveResults == null)
            {
                if (spans == null)
                {
                    return;
                }

                newSpans = spans;
            }
            else
            {
                if (spans == null)
                {
                    return;
                }

                // 列移動情報がある場合はJSONデータにあるインデックスとマッピングして、移動後のインデックスでセル結合を行う
                var spanList = new List<SpanDefinitions>();
                foreach (var span in spans)
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
                sheetView.AddColumnHeaderSpanCell(span.Row, span.Column, span.RowCount, span.ColumnCount);
            }
        }

        /// <summary>
        /// 列情報をコンパイルする。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="root">Root オブジェクト。</param>
        private static void CompileColumns(SheetView sheetView, SheetViewDefinitions root, TemplateColumnsDefinitions templateColumnsDefinitions)
        {
            var columns = root.Columns;
            if (columns == null)
            {
                throw new ArgumentNullException("Columns");
            }

            TemplateColumnDefinitions templateColumn = null;

            foreach (var column in columns.Select((Item, Index) => new { Item, Index }))
            {
                templateColumn = null;
                ICellType cellType;
                switch (column.Item.CellType.ToLower())
                {
                    //case "textcelltype":
                    //    cellType = new TextCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileTextCellType((TextCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "datetimecelltype":
                    //    cellType = new DateTimeCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileDateTimeCellType((DateTimeCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "numbercelltype":
                    //    cellType = new NumberCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileNumberCellType((NumberCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "comboboxcelltype":
                    //    cellType = new ComboBoxCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileComboBoxCellType((ComboBoxCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "checkboxcelltype":
                    //    cellType = new CheckBoxCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileCheckBoxCellType((CheckBoxCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "buttoncelltype":
                    //    cellType = new ButtonCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileButtonCellType((ButtonCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "gctextboxcelltype":
                    //    cellType = new GcTextBoxCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileGcTextBoxCellType((GcTextBoxCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "gcdatetimecelltype":
                    //    cellType = new GcDateTimeCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileGcDateTimeCellType((GcDateTimeCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "gcnumbercelltype":
                    //    cellType = new GcNumberCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    CompileGcNumberCellType((GcNumberCellType)cellType, column.Item.CellTypeProperties);
                    //    break;

                    //case "gccomboboxcelltype":
                    //    cellType = new GcComboBoxCellType();
                    //    sheetView.Columns[column.Index].CellType = cellType;
                    //    ((GcComboBoxCellType)cellType).DeserializeJson(Newtonsoft.Json.JsonConvert.SerializeObject(column.Item.CellTypeProps));
                    //    break;
                }

                if (string.Compare(nameof(GcNumberCellType), column.Item.CellType, true) == 0)
                {
                    cellType = new GcNumberCellType();
                    sheetView.Columns[column.Index].CellType = cellType;

                    // テンプレートで名前が見つかったら、それさき設定する
                    if (!string.IsNullOrEmpty(column.Item.TemplateName))
                    {
                        templateColumn = templateColumnsDefinitions.Columns.Where(x => x.TemplateName == column.Item.TemplateName).FirstOrDefault();
                        if (templateColumn != null)
                        {
                            if (templateColumn.CellTypeProps != null)
                            {
                                ((GcNumberCellType)cellType).DeserializeJson(Newtonsoft.Json.JsonConvert.SerializeObject(templateColumn.CellTypeProps));
                            }
                        }
                    }
                    if (column.Item.CellTypeProps != null)
                    {
                        ((GcNumberCellType)cellType).DeserializeJson(Newtonsoft.Json.JsonConvert.SerializeObject(column.Item.CellTypeProps));
                    }
                }

                if (string.Compare(nameof(GcComboBoxCellType), column.Item.CellType, true) == 0)
                {
                    cellType = new GcComboBoxCellType();
                    sheetView.Columns[column.Index].CellType = cellType;
                    ((GcComboBoxCellType)cellType).DeserializeJson(Newtonsoft.Json.JsonConvert.SerializeObject(column.Item.CellTypeProps));
                }


                // 列全体の定義を反映
                if (root.AllColumn != null)
                {
                    var allColumn = root.AllColumn;
                    if (allColumn.AllowAutoFilter.HasValue)
                    {
                        sheetView.Columns[column.Index].AllowAutoFilter = allColumn.AllowAutoFilter.Value;
                    }
                    if (allColumn.AllowAutoSort.HasValue)
                    {
                        sheetView.Columns[column.Index].AllowAutoSort = allColumn.AllowAutoSort.Value;
                    }
                    if (allColumn.HorizontalAlignment.HasValue)
                    {
                        sheetView.Columns[column.Index].HorizontalAlignment = allColumn.HorizontalAlignment.Value;
                    }
                    if (allColumn.ImeMode.HasValue)
                    {
                        sheetView.Columns[column.Index].ImeMode = allColumn.ImeMode.Value;
                    }
                    if (allColumn.Locked.HasValue)
                    {
                        sheetView.Columns[column.Index].Locked = allColumn.Locked.Value;
                    }
                    if (allColumn.VerticalAlignment.HasValue)
                    {
                        sheetView.Columns[column.Index].VerticalAlignment = allColumn.VerticalAlignment.Value;
                    }
                    if (allColumn.Visible.HasValue)
                    {
                        sheetView.Columns[column.Index].Visible = allColumn.Visible.Value;
                    }
                    if (allColumn.Width.HasValue)
                    {
                        sheetView.Columns[column.Index].Width = allColumn.Width.Value;
                    }
                }

                // テンプレート列の定義
                if (templateColumn != null)
                {
                    if (templateColumn.AllowAutoFilter.HasValue)
                    {
                        sheetView.Columns[column.Index].AllowAutoFilter = templateColumn.AllowAutoFilter.Value;
                    }
                    if (templateColumn.AllowAutoSort.HasValue)
                    {
                        sheetView.Columns[column.Index].AllowAutoSort = templateColumn.AllowAutoSort.Value;
                    }
                    if (!string.IsNullOrEmpty(templateColumn.DataField))
                    {
                        sheetView.BindDataColumn(column.Index, templateColumn.DataField);
                    }
                    if (templateColumn.HorizontalAlignment.HasValue)
                    {
                        sheetView.Columns[column.Index].HorizontalAlignment = templateColumn.HorizontalAlignment.Value;
                    }
                    if (templateColumn.ImeMode.HasValue)
                    {
                        sheetView.Columns[column.Index].ImeMode = templateColumn.ImeMode.Value;
                    }
                    if (templateColumn.Locked.HasValue)
                    {
                        sheetView.Columns[column.Index].Locked = templateColumn.Locked.Value;
                    }
                    if (templateColumn.VerticalAlignment.HasValue)
                    {
                        sheetView.Columns[column.Index].VerticalAlignment = templateColumn.VerticalAlignment.Value;
                    }
                    if (templateColumn.Visible.HasValue)
                    {
                        sheetView.Columns[column.Index].Visible = templateColumn.Visible.Value;
                    }
                    if (templateColumn.Width.HasValue)
                    {
                        sheetView.Columns[column.Index].Width = templateColumn.Width.Value;
                    }
                }

                // 列の定義
                if (column.Item.AllowAutoFilter.HasValue)
                {
                    sheetView.Columns[column.Index].AllowAutoFilter = column.Item.AllowAutoFilter.Value;
                }
                if (column.Item.AllowAutoSort.HasValue)
                {
                    sheetView.Columns[column.Index].AllowAutoSort = column.Item.AllowAutoSort.Value;
                }
                if (!string.IsNullOrEmpty(column.Item.DataField))
                {
                    sheetView.BindDataColumn(column.Index, column.Item.DataField);
                }
                if (column.Item.HorizontalAlignment.HasValue)
                {
                    sheetView.Columns[column.Index].HorizontalAlignment = column.Item.HorizontalAlignment.Value;
                }
                if (column.Item.ImeMode.HasValue)
                {
                    sheetView.Columns[column.Index].ImeMode = column.Item.ImeMode.Value;
                }
                if (column.Item.Locked.HasValue)
                {
                    sheetView.Columns[column.Index].Locked = column.Item.Locked.Value;
                }
                if (column.Item.VerticalAlignment.HasValue)
                {
                    sheetView.Columns[column.Index].VerticalAlignment = column.Item.VerticalAlignment.Value;
                }
                if (column.Item.Visible.HasValue)
                {
                    sheetView.Columns[column.Index].Visible = column.Item.Visible.Value;
                }
                if (column.Item.Width.HasValue)
                {
                    sheetView.Columns[column.Index].Width = column.Item.Width.Value;
                }
            }
        }

        /// <summary>
        /// TextCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">TextCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileTextCellType(TextCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();
                if (propKey == "maxlength")
                {
                    cellType.MaxLength = Convert.ToInt32(prop.Value);
                }
                if (propKey == "charactercasing")
                {
                    CharacterCasing characterCasing;
                    if (EnumExt.TryParse(prop.Value.ToString(), out characterCasing))
                    {
                        cellType.CharacterCasing = characterCasing;
                    }
                }
                if (propKey == "characterset")
                {
                    CharacterSet characterSet;
                    if (EnumExt.TryParse(prop.Value.ToString(), out characterSet))
                    {
                        cellType.CharacterSet = characterSet;
                    }
                }
                if (propKey == "multiline")
                {
                    cellType.Multiline = (bool)prop.Value;
                }
                if (propKey == "readonly")
                {
                    cellType.ReadOnly = (bool)prop.Value;
                }
                if (propKey == "static")
                {
                    cellType.Static = (bool)prop.Value;
                }
                if (propKey == "enablesubeditor")
                {
                    cellType.EnableSubEditor = (bool)prop.Value;
                }
            }
        }

        /// <summary>
        /// DateTimeCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">DateTimeCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileDateTimeCellType(DateTimeCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();
                if (propKey == "calendar")
                {
                    var propValue = prop.Value.ToString().ToLower();
                    if (propValue == "gregorianlocalized")
                    {
                        cellType.Calendar = new GregorianCalendar(GregorianCalendarTypes.Localized);
                    }
                    if (propValue == "japanese")
                    {
                        cellType.Calendar = new JapaneseCalendar();
                    }
                    if (propValue == "gregorianus")
                    {
                        cellType.Calendar = new GregorianCalendar(GregorianCalendarTypes.USEnglish);
                    }
                }
                if (propKey == "datetimeformat")
                {
                    DateTimeFormat dateTimeFormat;
                    if (EnumExt.TryParse(prop.Value.ToString(), out dateTimeFormat))
                    {
                        cellType.DateTimeFormat = dateTimeFormat;
                    }
                }
                if (propKey == "maximumdate")
                {
                    cellType.MaximumDate = DateTime.Parse(prop.Value.ToString());
                }
                if (propKey == "maximumtime")
                {
                    var date = DateTime.Parse(prop.Value.ToString());
                    var timeSpan = new TimeSpan(new TimeSpan(0, date.Hour, date.Minute, date.Second, 999).Ticks + 9999);
                    cellType.MaximumTime = timeSpan;
                }
                if (propKey == "minumumdate")
                {
                    cellType.MinimumDate = DateTime.Parse(prop.Value.ToString());
                }
                if (propKey == "minumumtime")
                {
                    var date = DateTime.Parse(prop.Value.ToString());
                    var timeSpan = new TimeSpan(0, date.Hour, date.Minute, date.Second);
                    cellType.MinimumTime = timeSpan;
                }
                if (propKey == "readonly")
                {
                    cellType.ReadOnly = (bool)prop.Value;
                }
                if (propKey == "static")
                {
                    cellType.Static = (bool)prop.Value;
                }
                if (propKey == "enablesubeditor")
                {
                    cellType.EnableSubEditor = (bool)prop.Value;
                }
            }
        }

        /// <summary>
        /// NumberCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">NumberCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileNumberCellType(NumberCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();
                if (propKey == "decimalplaces")
                {
                    cellType.DecimalPlaces = Convert.ToInt32(prop.Value);
                }
                if (propKey == "decimalseparator")
                {
                    cellType.DecimalSeparator = (string)prop.Value;
                }
                if (propKey == "fixedpoint")
                {
                    cellType.FixedPoint = (bool)prop.Value;
                }
                if (propKey == "leadingzero")
                {
                    LeadingZero leadingZero;
                    if (EnumExt.TryParse(prop.Value.ToString(), out leadingZero))
                    {
                        cellType.LeadingZero = leadingZero;
                    }
                }
                if (propKey == "maximumvalue")
                {
                    cellType.MaximumValue = Convert.ToDouble(prop.Value);
                }
                if (propKey == "minimumvalue")
                {
                    cellType.MinimumValue = Convert.ToDouble(prop.Value);
                }
                if (propKey == "negativeformat")
                {
                    NegativeFormat negativeFormat;
                    if (EnumExt.TryParse(prop.Value.ToString(), out negativeFormat))
                    {
                        cellType.NegativeFormat = negativeFormat;
                    }
                }
                if (propKey == "negativered")
                {
                    cellType.NegativeRed = (bool)prop.Value;
                }
                if (propKey == "readonly")
                {
                    cellType.ReadOnly = (bool)prop.Value;
                }
                if (propKey == "separator")
                {
                    cellType.Separator = (string)prop.Value;
                }
                if (propKey == "showseparator")
                {
                    cellType.ShowSeparator = (bool)prop.Value;
                }
                if (propKey == "static")
                {
                    cellType.Static = (bool)prop.Value;
                }
                if (propKey == "enablesubeditor")
                {
                    cellType.EnableSubEditor = (bool)prop.Value;
                }
            }
        }

        /// <summary>
        /// ComboBoxCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">ComboBoxCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileComboBoxCellType(ComboBoxCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();
                if (propKey == "maxdrop")
                {
                    cellType.MaxDrop = Convert.ToInt32(prop.Value);
                }
                if (propKey == "editable")
                {
                    cellType.Editable = (bool)prop.Value;
                }
                if (propKey == "editorvalue")
                {
                    EditorValue editorValue;
                    if (EnumExt.TryParse(prop.Value.ToString(), out editorValue))
                    {
                        cellType.EditorValue = editorValue;
                    }
                }
            }
        }

        /// <summary>
        /// CheckBoxCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">CheckBoxCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileCheckBoxCellType(CheckBoxCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();
                if (propKey == "texttrue")
                {
                    cellType.TextTrue = (string)prop.Value;
                }
                if (propKey == "textfalse")
                {
                    cellType.TextFalse = (string)prop.Value;
                }
                if (propKey == "textindeterminate")
                {
                    cellType.TextIndeterminate = (string)prop.Value;
                }
                if (propKey == "threestate")
                {
                    cellType.ThreeState = (bool)prop.Value;
                }
            }
        }

        /// <summary>
        /// ButtonCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">ButtonCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileButtonCellType(ButtonCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();
                if (propKey == "text")
                {
                    cellType.Text = (string)prop.Value;
                }
                if (propKey == "textalign")
                {
                    ButtonTextAlign textAlign;
                    if (EnumExt.TryParse(prop.Value.ToString(), out textAlign))
                    {
                        cellType.TextAlign = textAlign;
                    }
                }
                if (propKey == "twostate")
                {
                    cellType.TwoState = (bool)prop.Value;
                }
                if (propKey == "textdown")
                {
                    cellType.TextDown = (string)prop.Value;
                }
                if (propKey == "wordwrap")
                {
                    cellType.WordWrap = (bool)prop.Value;
                }
            }
        }

        /// <summary>
        /// GcTextBoxCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">GcTextBoxCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileGcTextBoxCellType(GcTextBoxCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();

                if (propKey == "acceptscrlf")
                {
                    CrLfMode crLfMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out crLfMode))
                    {
                        cellType.AcceptsCrLf = crLfMode;
                    }
                }
                if (propKey == "acceptstabchar")
                {
                    TabCharMode tabCharMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out tabCharMode))
                    {
                        cellType.AcceptsTabChar = tabCharMode;
                    }
                }
                if (propKey == "allowspace")
                {
                    AllowSpace allowSpace;
                    if (EnumExt.TryParse(prop.Value.ToString(), out allowSpace))
                    {
                        cellType.AllowSpace = allowSpace;
                    }
                }
                if (propKey == "autoconvert")
                {
                    cellType.AutoConvert = (bool)prop.Value;
                }
                if (propKey == "editmode")
                {
                    EditMode editMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out editMode))
                    {
                        cellType.EditMode = editMode;
                    }
                }
                if (propKey == "ellipsis")
                {
                    EllipsisMode ellipsisMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out ellipsisMode))
                    {
                        cellType.Ellipsis = ellipsisMode;
                    }
                }
                if (propKey == "ellipsisstring")
                {
                    cellType.EllipsisString = (string)prop.Value;
                }
                if (propKey == "excelexportformat")
                {
                    cellType.ExcelExportFormat = (string)prop.Value;
                }
                if (propKey == "exitonlastchar")
                {
                    cellType.ExitOnLastChar = (bool)prop.Value;
                }
                if (propKey == "focusposition")
                {
                    EditorBaseFocusCursorPosition editorBaseFocusCursorPosition;
                    if (EnumExt.TryParse(prop.Value.ToString(), out editorBaseFocusCursorPosition))
                    {
                        cellType.FocusPosition = editorBaseFocusCursorPosition;
                    }
                }
                if (propKey == "formatstring")
                {
                    cellType.FormatString = (string)prop.Value;
                }
                if (propKey == "linespace")
                {
                    cellType.LineSpace = Convert.ToInt32(prop.Value);
                }
                if (propKey == "maxlength")
                {
                    cellType.MaxLength = Convert.ToInt32(prop.Value);
                }
                if (propKey == "maxlengthcodepage")
                {
                    cellType.MaxLengthCodePage = Convert.ToInt32(prop.Value);
                }
                if (propKey == "maxlengthunit")
                {
                    LengthUnit lengthUnit;
                    if (EnumExt.TryParse(prop.Value.ToString(), out lengthUnit))
                    {
                        cellType.MaxLengthUnit = lengthUnit;
                    }
                }
                if (propKey == "maxlinecount")
                {
                    cellType.MaxLineCount = Convert.ToInt32(prop.Value);
                }
                if (propKey == "multiline")
                {
                    cellType.Multiline = (bool)prop.Value;
                }
                if (propKey == "passwordchar")
                {
                    cellType.PasswordChar = Convert.ToChar(prop.Value);
                }
                if (propKey == "passwordrevelationmode")
                {
                    PasswordRevelationMode passwordRevelationMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out passwordRevelationMode))
                    {
                        cellType.PasswordRevelationMode = passwordRevelationMode;
                    }
                }
                if (propKey == "readonly")
                {
                    cellType.ReadOnly = (bool)prop.Value;
                }
                if (propKey == "recommendedvalue")
                {
                    cellType.RecommendedValue = (string)prop.Value;
                }
                if (propKey == "showrecommendedvalue")
                {
                    cellType.ShowRecommendedValue = (bool)prop.Value;
                }
                if (propKey == "static")
                {
                    cellType.Static = (bool)prop.Value;
                }
                if (propKey == "usesystempasswordchar")
                {
                    cellType.UseSystemPasswordChar = (bool)prop.Value;
                }
                if (propKey == "wrapmode")
                {
                    WrapMode wrapMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out wrapMode))
                    {
                        cellType.WrapMode = wrapMode;
                    }
                }
            }
        }

        /// <summary>
        /// GcDateTimeCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">GcDateTimeCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileGcDateTimeCellType(GcDateTimeCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();

                if (propKey == "acceptscrlf")
                {
                    CrLfMode crLfMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out crLfMode))
                    {
                        cellType.AcceptsCrLf = crLfMode;
                    }
                }
                if (propKey == "clipcontent")
                {
                    ClipContent clipContent;
                    if (EnumExt.TryParse(prop.Value.ToString(), out clipContent))
                    {
                        cellType.ClipContent = clipContent;
                    }
                }
                if (propKey == "defaultactivefield")
                {
                    cellType.SerializationDefaultActiveFieldIndex = Convert.ToInt32(prop.Value);
                }
                if (propKey == "editmode")
                {
                    EditMode editMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out editMode))
                    {
                        cellType.EditMode = editMode;
                    }
                }
                if (propKey == "displayfields")
                {
                    cellType.DisplayFields.Clear();
                    if (prop.Value is JArray)
                    {
                        var list = ((JArray)prop.Value).ToList();
                        DateDisplayFieldInfo info;
                        foreach (var item in list)
                        {
                            switch (item["Type"].ToString().ToLower())
                            {
                                case "dateyeardisplayfieldinfo":
                                    info = item.ToObject<DateYearDisplayFieldInfo>();
                                    cellType.DisplayFields.Add(info);
                                    break;
                                case "datemonthdisplayfieldinfo":
                                    info = item.ToObject<DateMonthDisplayFieldInfo>();
                                    cellType.DisplayFields.Add(info);
                                    break;
                                case "datedaydisplayfieldinfo":
                                    info = item.ToObject<DateDayDisplayFieldInfo>();
                                    cellType.DisplayFields.Add(info);
                                    break;
                                case "datehourdisplayfieldinfo":
                                    info = item.ToObject<DateHourDisplayFieldInfo>();
                                    cellType.DisplayFields.Add(info);
                                    break;
                                case "dateminutedisplayfieldinfo":
                                    info = item.ToObject<DateMinuteDisplayFieldInfo>();
                                    cellType.DisplayFields.Add(info);
                                    break;
                                case "dateseconddisplayfieldinfo":
                                    info = item.ToObject<DateSecondDisplayFieldInfo>();
                                    cellType.DisplayFields.Add(info);
                                    break;
                                case "dateliteraldisplayfieldinfo":
                                    info = item.ToObject<DateLiteralDisplayFieldInfo>();
                                    cellType.DisplayFields.Add(info);
                                    break;
                            }
                        }
                    }
                    else
                    {
                        cellType.DisplayFields.AddRange((string)prop.Value);
                    }
                }
                if (propKey == "excelexportformat")
                {
                    cellType.ExcelExportFormat = (string)prop.Value;
                }
                if (propKey == "exitonlastchar")
                {
                    cellType.ExitOnLastChar = (bool)prop.Value;
                }
                if (propKey == "fields")
                {
                    cellType.Fields.Clear();
                    if (prop.Value is JArray)
                    {
                        var list = ((JArray)prop.Value).ToList();
                        DateFieldInfo info;
                        foreach (var item in list)
                        {
                            switch (item["Type"].ToString().ToLower())
                            {
                                case "dateyearfieldinfo":
                                    info = item.ToObject<DateYearFieldInfo>();
                                    cellType.Fields.Add(info);
                                    break;
                                case "datemonthfieldinfo":
                                    info = item.ToObject<DateMonthFieldInfo>();
                                    cellType.Fields.Add(info);
                                    break;
                                case "datedayfieldinfo":
                                    info = item.ToObject<DateDayFieldInfo>();
                                    cellType.Fields.Add(info);
                                    break;
                                case "datehourfieldinfo":
                                    info = item.ToObject<DateHourFieldInfo>();
                                    cellType.Fields.Add(info);
                                    break;
                                case "dateminutefieldinfo":
                                    info = item.ToObject<DateMinuteFieldInfo>();
                                    cellType.Fields.Add(info);
                                    break;
                                case "datesecondfieldinfo":
                                    info = item.ToObject<DateSecondFieldInfo>();
                                    cellType.Fields.Add(info);
                                    break;
                                case "dateliteralfieldinfo":
                                    info = item.ToObject<DateLiteralFieldInfo>();
                                    cellType.Fields.Add(info);
                                    break;
                            }
                        }

                    }
                    else
                    {
                        cellType.Fields.AddRange((string)prop.Value);
                    }
                }
                if (propKey == "fieldseditmode")
                {
                    FieldsEditMode fieldsEditMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out fieldsEditMode))
                    {
                        cellType.FieldsEditMode = fieldsEditMode;
                    }
                }
                if (propKey == "focusposition")
                {
                    FieldsEditorFocusCursorPosition fieldsEditorFocusCursorPosition;
                    if (EnumExt.TryParse(prop.Value.ToString(), out fieldsEditorFocusCursorPosition))
                    {
                        cellType.FocusPosition = fieldsEditorFocusCursorPosition;
                    }
                }
                if (propKey == "maxdate")
                {
                    cellType.MaxDate = Convert.ToDateTime(prop.Value);
                }
                if (propKey == "maxminbehavior")
                {
                    MaxMinBehavior maxMinBehavior;
                    if (EnumExt.TryParse(prop.Value.ToString(), out maxMinBehavior))
                    {
                        cellType.MaxMinBehavior = maxMinBehavior;
                    }
                }
                if (propKey == "promptchar")
                {
                    cellType.PromptChar = Convert.ToChar(prop.Value);
                }
                if (propKey == "readonly")
                {
                    cellType.ReadOnly = (bool)prop.Value;
                }
                if (propKey == "recommendedvalue")
                {
                    if (prop.Value == null)
                    {
                        cellType.RecommendedValue = null;
                        continue;
                    }
                    cellType.RecommendedValue = Convert.ToDateTime(prop.Value);
                }
                if (propKey == "showliterals")
                {
                    ShowLiterals showLiterals;
                    if (EnumExt.TryParse(prop.Value.ToString(), out showLiterals))
                    {
                        cellType.ShowLiterals = showLiterals;
                    }
                }
                if (propKey == "showrecommendedvalue")
                {
                    cellType.ShowRecommendedValue = (bool)prop.Value;
                }
                if (propKey == "static")
                {
                    cellType.Static = (bool)prop.Value;
                }
                if (propKey == "tabaction")
                {
                    TabAction tabAction;
                    if (EnumExt.TryParse(prop.Value.ToString(), out tabAction))
                    {
                        cellType.TabAction = tabAction;
                    }
                }
                if (propKey == "validatemode")
                {
                    ValidateModeEx validateModeEx;
                    if (EnumExt.TryParse(prop.Value.ToString(), out validateModeEx))
                    {
                        cellType.ValidateMode = validateModeEx;
                    }
                }
                if (propKey == "dropdownallowdrop")
                {
                    cellType.DropDown.AllowDrop = (bool)prop.Value;
                }
                if (propKey == "sidebuttonsclear")
                {
                    if ((bool)prop.Value)
                    {
                        cellType.SideButtons.Clear();
                    }
                }
            }
        }

        /// <summary>
        /// GcNumberCellType をコンパイルする。
        /// </summary>
        /// <param name="cellType">GcNumberCellType オブジェクト。</param>
        /// <param name="cellTypeProperties">セルタイプ情報。</param>
        private static void CompileGcNumberCellType(GcNumberCellType cellType, Dictionary<string, object> cellTypeProperties)
        {
            if (cellTypeProperties == null)
            {
                return;
            }

            foreach (var prop in cellTypeProperties)
            {
                var propKey = prop.Key.ToLower();

                if (propKey == "acceptscrlf")
                {
                    CrLfMode crLfMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out crLfMode))
                    {
                        cellType.AcceptsCrLf = crLfMode;
                    }
                }
                if (propKey == "acceptsdecimal")
                {
                    DecimalMode decimalMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out decimalMode))
                    {
                        cellType.AcceptsDecimal = decimalMode;
                    }
                }
                if (propKey == "allowdeletetonull")
                {
                    cellType.AllowDeleteToNull = (bool)prop.Value;
                }
                if (propKey == "clipcontent")
                {
                    ClipContent clipContent;
                    if (EnumExt.TryParse(prop.Value.ToString(), out clipContent))
                    {
                        cellType.ClipContent = clipContent;
                    }
                }

                if (propKey == "displayfields")
                {
                    cellType.DisplayFields.Clear();
                    var list = ((JArray)prop.Value).ToList();
                    NumberDisplayFieldInfo info;
                    foreach (var item in list)
                    {
                        switch (item["Type"].ToString().ToLower())
                        {
                            case "numbersigndisplayfieldinfo":
                                info = item.ToObject<NumberSignDisplayFieldInfo>();
                                cellType.DisplayFields.Add(info);
                                break;
                            case "numberintegerpartdisplayfieldinfo":
                                info = item.ToObject<NumberIntegerPartDisplayFieldInfo>();
                                cellType.DisplayFields.Add(info);
                                break;
                            case "numberdecimalseparatordisplayfieldinfo":
                                info = item.ToObject<NumberDecimalSeparatorDisplayFieldInfo>();
                                cellType.DisplayFields.Add(info);
                                break;
                            case "numberdecimalpartdisplayfieldinfo":
                                info = item.ToObject<NumberDecimalPartDisplayFieldInfo>();
                                cellType.DisplayFields.Add(info);
                                break;
                        }
                    }
                }

                if (propKey == "editmode")
                {
                    EditMode editMode;
                    if (EnumExt.TryParse(prop.Value.ToString(), out editMode))
                    {
                        cellType.EditMode = editMode;
                    }
                }
                if (propKey == "excelexportformat")
                {
                    cellType.ExcelExportFormat = (string)prop.Value;
                }
                if (propKey == "exitonlastchar")
                {
                    cellType.ExitOnLastChar = (bool)prop.Value;
                }

                if (propKey == "fields")
                {
                    var list = ((JArray)prop.Value).ToList();

                    foreach (var field in cellType.Fields.Select((Item, Index) => new { Item, Index }))
                    {
                        switch (field.Index)
                        {
                            case 0:
                            case 4:
                                var itemExpression = list.Where(x => x["Type"].ToString().ToLower() == "numbersignfieldinfo").Select((Item, Index) => new { Item, Index });

                                NumberSignFieldInfo signInfo;
                                NumberSignFieldInfo signSetInfo;
                                if (field.Index == 0)
                                {
                                    // NOTE: NumberSignFieldInfo の定義が1つある場合は必ず最初の NumberSignFieldInfo とみなす。
                                    signInfo = cellType.Fields[0] as NumberSignFieldInfo;
                                    signSetInfo = itemExpression.FirstOrDefault()?.Item.ToObject<NumberSignFieldInfo>();
                                }
                                else
                                {
                                    // NOTE: NumberSignFieldInfo の定義が2つ以上ない場合は制御せず、必ず最後の NumberSignFieldInfo とみなす。
                                    if (itemExpression.Count() < 2)
                                    {
                                        continue;
                                    }
                                    signInfo = cellType.Fields[4] as NumberSignFieldInfo;
                                    signSetInfo = itemExpression.LastOrDefault()?.Item.ToObject<NumberSignFieldInfo>();
                                }
                                if (signSetInfo == null)
                                {
                                    continue;
                                }
                                signInfo.BackColor = signSetInfo.BackColor;
                                signInfo.Font = signSetInfo.Font;
                                signInfo.ForeColor = signSetInfo.ForeColor;
                                signInfo.Margin = signSetInfo.Margin;
                                signInfo.NegativePattern = signSetInfo.NegativePattern;
                                signInfo.Padding = signSetInfo.Padding;
                                signInfo.PositivePattern = signSetInfo.PositivePattern;
                                signInfo.Text = signSetInfo.Text;
                                break;

                            case 1:
                                var intPartSetInfoItem = list.Where(x => x["Type"].ToString().ToLower() == "numberintegerpartfieldinfo").FirstOrDefault();
                                if (intPartSetInfoItem == null)
                                {
                                    continue;
                                }
                                var intPartInfo = cellType.Fields[1] as NumberIntegerPartFieldInfo;
                                var intPartSetInfo = intPartSetInfoItem.ToObject<NumberIntegerPartFieldInfo>();
                                intPartInfo.BackColor = intPartSetInfo.BackColor;
                                intPartInfo.Font = intPartSetInfo.Font;
                                intPartInfo.ForeColor = intPartSetInfo.ForeColor;
                                intPartInfo.GroupSeparator = intPartSetInfo.GroupSeparator;
                                intPartInfo.GroupSizes = intPartSetInfo.GroupSizes;
                                intPartInfo.Margin = intPartSetInfo.Margin;
                                intPartInfo.MaxDigits = intPartSetInfo.MaxDigits;
                                intPartInfo.MinDigits = intPartSetInfo.MinDigits;
                                intPartInfo.Padding = intPartSetInfo.Padding;
                                intPartInfo.SpinIncrement = intPartSetInfo.SpinIncrement;
                                intPartInfo.Text = intPartSetInfo.Text;
                                break;

                            case 2:
                                var decSepSetInfoItem = list.Where(x => x["Type"].ToString().ToLower() == "numberdecimalseparatorfieldinfo").FirstOrDefault();
                                if (decSepSetInfoItem == null)
                                {
                                    continue;
                                }
                                var decSepInfo = cellType.Fields[2] as NumberDecimalSeparatorFieldInfo;
                                var decSepSetInfo = decSepSetInfoItem.ToObject<NumberDecimalSeparatorFieldInfo>();
                                decSepInfo.BackColor = decSepSetInfo.BackColor;
                                decSepInfo.DecimalSeparator = decSepSetInfo.DecimalSeparator;
                                decSepInfo.Font = decSepSetInfo.Font;
                                decSepInfo.ForeColor = decSepSetInfo.ForeColor;
                                decSepInfo.Margin = decSepSetInfo.Margin;
                                decSepInfo.Padding = decSepSetInfo.Padding;
                                decSepInfo.Text = decSepSetInfo.Text;
                                break;

                            case 3:
                                var decPartSetInfoItem = list.Where(x => x["Type"].ToString().ToLower() == "numberdecimalpartfieldinfo").FirstOrDefault();
                                var decPartInfo = cellType.Fields[3] as NumberDecimalPartFieldInfo;
                                var decPartSetInfo = decPartSetInfoItem.ToObject<NumberDecimalPartFieldInfo>();
                                decPartInfo.BackColor = decPartSetInfo.BackColor;
                                decPartInfo.Font = decPartSetInfo.Font;
                                decPartInfo.ForeColor = decPartSetInfo.ForeColor;
                                decPartInfo.Margin = decPartSetInfo.Margin;
                                decPartInfo.MaxDigits = decPartSetInfo.MaxDigits;
                                decPartInfo.MinDigits = decPartSetInfo.MinDigits;
                                decPartInfo.Padding = decPartSetInfo.Padding;
                                decPartInfo.SpinIncrement = decPartSetInfo.SpinIncrement;
                                decPartInfo.Text = decPartSetInfo.Text;
                                break;
                        }
                    }
                }

                if (propKey == "focusposition")
                {
                    EditorBaseFocusCursorPosition editorBaseFocusCursorPosition;
                    if (EnumExt.TryParse(prop.Value.ToString(), out editorBaseFocusCursorPosition))
                    {
                        cellType.FocusPosition = editorBaseFocusCursorPosition;
                    }
                }
                if (propKey == "maxminbehavior")
                {
                    MaxMinBehavior maxMinBehavior;
                    if (EnumExt.TryParse(prop.Value.ToString(), out maxMinBehavior))
                    {
                        cellType.MaxMinBehavior = maxMinBehavior;
                    }
                }
                if (propKey == "maxvalue")
                {
                    cellType.MaxValue = Convert.ToDecimal(prop.Value);
                }
                if (propKey == "minvalue")
                {
                    cellType.MinValue = Convert.ToDecimal(prop.Value);
                }
                if (propKey == "negativecolor")
                {
                    cellType.NegativeColor = ColorTranslator.FromHtml((string)prop.Value);
                }
                if (propKey == "readonly")
                {
                    cellType.ReadOnly = (bool)prop.Value;
                }
                if (propKey == "recommendedvalue")
                {
                    cellType.RecommendedValue = Convert.ToDecimal(prop.Value);
                }
                if (propKey == "roundpattern")
                {
                    RoundPattern roundPattern;
                    if (EnumExt.TryParse(prop.Value.ToString(), out roundPattern))
                    {
                        cellType.RoundPattern = roundPattern;
                    }
                }
                if (propKey == "showrecommendedvalue")
                {
                    cellType.ShowRecommendedValue = (bool)prop.Value;
                }
                if (propKey == "static")
                {
                    cellType.Static = (bool)prop.Value;
                }
                if (propKey == "usenegativecolor")
                {
                    cellType.UseNegativeColor = (bool)prop.Value;
                }
                if (propKey == "valuesign")
                {
                    ValueSignControl valueSignControl;
                    if (EnumExt.TryParse(prop.Value.ToString(), out valueSignControl))
                    {
                        cellType.ValueSign = valueSignControl;
                    }
                }
                if (propKey == "dropdownallowdrop")
                {
                    cellType.DropDown.AllowDrop = (bool)prop.Value;
                }
                if (propKey == "sidebuttonsclear")
                {
                    if ((bool)prop.Value)
                    {
                        cellType.SideButtons.Clear();
                    }
                }
            }
        }

        /// <summary>
        /// 現在のシートからヘッダー行/列情報を逆コンパイルする。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <returns>HeaderRowJson[] オブジェクト。</returns>
        private static HeaderRowDefinitions[] DecompileColumnHeaderRows(SheetView sheetView)
        {
            var headerRows = new List<HeaderRowDefinitions>();
            foreach (Row row in sheetView.ColumnHeader.Rows)
            {
                var headerRow = new HeaderRowDefinitions()
                {
                    Height = row.Height,
                    HorizontalAlignment = row.HorizontalAlignment,
                    VerticalAlignment = row.VerticalAlignment
                };

                var headerColumns = new List<HeaderColumnDefinitions>();
                foreach (Column column in sheetView.ColumnHeader.Columns)
                {
                    var headerColumn = new HeaderColumnDefinitions();
                    headerColumn.Value = sheetView.ColumnHeader.Cells[row.Index, column.Index].Value?.ToString();
                    headerColumn.HorizontalAlignment = sheetView.ColumnHeader.Cells[row.Index, column.Index].HorizontalAlignment;
                    headerColumn.VerticalAlignment = sheetView.ColumnHeader.Cells[row.Index, column.Index].VerticalAlignment;
                    headerColumns.Add(headerColumn);
                }
                headerRow.Columns = headerColumns.ToArray();
                headerRows.Add(headerRow);
            }

            return headerRows.ToArray();
        }

        /// <summary>
        /// 現在のシートからヘッダースパン情報を逆コンパイルする。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <returns>SpanJson[] オブジェクト。</returns>
        private static SpanDefinitions[] DecompileColumnHeaderSpans(SheetView sheetView)
        {
            var headerSpans = new List<SpanDefinitions>();
            foreach (Row row in sheetView.ColumnHeader.Rows)
            {
                foreach (Column column in sheetView.ColumnHeader.Columns)
                {
                    var spanRange = sheetView.GetColumnHeaderSpanCell(row.Index, column.Index);
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
        /// 現在のシートから列情報を逆コンパイルする。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="columnOptions">列のオプション情報設定アクションデリゲート。</param>
        /// <returns>ColumnJson[] オブジェクト</returns>
        private static ColumnDefinitions[] DecompileColumns(SheetView sheetView, Func<Column, Dictionary<string, object>> columnOptions)
        {
            var columns = new List<ColumnDefinitions>();

            foreach (Column svColumn in sheetView.Columns)
            {
                // TODO: GcComboBoxのみ動くようにする
                if (!(svColumn.CellType is GcComboBoxCellType) && !(svColumn.CellType is GcNumberCellType))
                {
                    continue;
                }


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

                if (svColumn.CellType is TextCellType)
                {
                    column.CellTypeProps = DecompileTextCellTypeProperties((TextCellType)svColumn.CellType);
                }
                if (svColumn.CellType is DateTimeCellType)
                {
                    column.CellTypeProps = DecompileDateTimeCellTypeProperties((DateTimeCellType)svColumn.CellType);
                }
                if (svColumn.CellType is NumberCellType)
                {
                    column.CellTypeProps = DecompileNumberCellTypeProperties((NumberCellType)svColumn.CellType);
                }
                if (svColumn.CellType is ComboBoxCellType)
                {
                    column.CellTypeProps = DecompileComboBoxCellTypeProperties((ComboBoxCellType)svColumn.CellType);
                }
                if (svColumn.CellType is CheckBoxCellType)
                {
                    column.CellTypeProps = DecompileCheckBoxCellTypeProperties((CheckBoxCellType)svColumn.CellType);
                }
                if (svColumn.CellType is ButtonCellType)
                {
                    column.CellTypeProps = DecompileButtonCellTypeProperties((ButtonCellType)svColumn.CellType);
                }
                if (svColumn.CellType is GcTextBoxCellType)
                {
                    column.CellTypeProps = DecompileGcTextBoxCellTypeProperties((GcTextBoxCellType)svColumn.CellType);
                }
                if (svColumn.CellType is GcDateTimeCellType)
                {
                    column.CellTypeProps = DecompileGcDateTimeCellTypeProperties((GcDateTimeCellType)svColumn.CellType);
                }
                if (svColumn.CellType is GcNumberCellType)
                {
                    //column.CellTypeProps = DecompileGcNumberCellTypeProperties((GcNumberCellType)svColumn.CellType);
                    column.CellTypeProps = JObject.Parse(((GcNumberCellType)svColumn.CellType).SerializeJson());
                }
                if (svColumn.CellType is GcComboBoxCellType)
                {
                    column.CellTypeProps = JObject.Parse(((GcComboBoxCellType)svColumn.CellType).SerializeJson());
                }

                column.Options = columnOptions?.Invoke(svColumn);

                columns.Add(column);
            }

            return columns.ToArray();
        }

        /// <summary>
        /// TextCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">TextCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileTextCellTypeProperties(TextCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.MaxLength), cellType.MaxLength);
            result.Add(nameof(cellType.CharacterCasing), cellType.CharacterCasing);
            result.Add(nameof(cellType.CharacterSet), cellType.CharacterSet);
            result.Add(nameof(cellType.Multiline), cellType.Multiline);
            result.Add(nameof(cellType.ReadOnly), cellType.ReadOnly);
            result.Add(nameof(cellType.Static), cellType.Static);
            result.Add(nameof(cellType.EnableSubEditor), cellType.EnableSubEditor);

            return result;
        }

        /// <summary>
        /// DateTieCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">DateTieCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileDateTimeCellTypeProperties(DateTimeCellType cellType)
        {
            var result = new Dictionary<string, object>();

            var calendarType = cellType.Calendar.GetType();
            if (calendarType.Name == "GregorianCalendar")
            {
                var pi = calendarType.GetProperty("CalendarType");
                var calendarTypeValue = pi.GetValue(cellType.Calendar).ToString();
                if (calendarTypeValue == "Localized")
                {
                    result.Add(nameof(cellType.Calendar), "GregorianLocalized");
                }
                else
                {
                    result.Add(nameof(cellType.Calendar), "GregorianUS");
                }

            }
            if (calendarType.Name == "JapaneseCalendar")
            {
                result.Add(nameof(cellType.Calendar), "Japanese");
            }

            result.Add(nameof(cellType.DateTimeFormat), cellType.DateTimeFormat);
            result.Add(nameof(cellType.MaximumDate), cellType.MaximumDate.ToShortDateString());
            result.Add(nameof(cellType.MaximumTime), cellType.MaximumTime.ToString("hh':'mm':'ss"));
            result.Add(nameof(cellType.MinimumDate), cellType.MinimumDate.ToShortDateString());
            result.Add(nameof(cellType.MinimumTime), cellType.MinimumTime.ToString("hh':'mm':'ss"));
            result.Add(nameof(cellType.ReadOnly), cellType.ReadOnly);
            result.Add(nameof(cellType.Static), cellType.Static);
            result.Add(nameof(cellType.EnableSubEditor), cellType.EnableSubEditor);

            return result;
        }

        /// <summary>
        /// NumberCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">NumberCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileNumberCellTypeProperties(NumberCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.DecimalPlaces), cellType.DecimalPlaces);
            result.Add(nameof(cellType.DecimalSeparator), cellType.DecimalSeparator);
            result.Add(nameof(cellType.FixedPoint), cellType.FixedPoint);
            result.Add(nameof(cellType.LeadingZero), cellType.LeadingZero);
            result.Add(nameof(cellType.MaximumValue), cellType.MaximumValue);
            result.Add(nameof(cellType.MinimumValue), cellType.MinimumValue);
            result.Add(nameof(cellType.NegativeFormat), cellType.NegativeFormat);
            result.Add(nameof(cellType.NegativeRed), cellType.NegativeRed);
            result.Add(nameof(cellType.ReadOnly), cellType.ReadOnly);
            result.Add(nameof(cellType.Separator), cellType.Separator);
            result.Add(nameof(cellType.ShowSeparator), cellType.ShowSeparator);
            result.Add(nameof(cellType.Static), cellType.Static);
            result.Add(nameof(cellType.EnableSubEditor), cellType.EnableSubEditor);

            return result;
        }

        /// <summary>
        /// ComboBoxCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">ComboBoxCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileComboBoxCellTypeProperties(ComboBoxCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.MaxDrop), cellType.MaxDrop);
            result.Add(nameof(cellType.Editable), cellType.Editable);
            result.Add(nameof(cellType.EditorValue), cellType.EditorValue);

            return result;
        }

        /// <summary>
        /// CheckCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">CheckBoxCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileCheckBoxCellTypeProperties(CheckBoxCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.TextTrue), cellType.TextTrue);
            result.Add(nameof(cellType.TextFalse), cellType.TextFalse);
            result.Add(nameof(cellType.TextIndeterminate), cellType.TextIndeterminate);
            result.Add(nameof(cellType.ThreeState), cellType.ThreeState);

            return result;
        }

        /// <summary>
        /// ButtonCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">ButtonCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileButtonCellTypeProperties(ButtonCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.Text), cellType.Text);
            result.Add(nameof(cellType.TextAlign), cellType.TextAlign);
            result.Add(nameof(cellType.TwoState), cellType.TwoState);
            result.Add(nameof(cellType.TextDown), cellType.TextDown);
            result.Add(nameof(cellType.WordWrap), cellType.WordWrap);

            return result;
        }

        /// <summary>
        /// GcTextBoxCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">GcTextBoxCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileGcTextBoxCellTypeProperties(GcTextBoxCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.AcceptsCrLf), cellType.AcceptsCrLf);
            result.Add(nameof(cellType.AcceptsTabChar), cellType.AcceptsTabChar);
            result.Add(nameof(cellType.AllowSpace), cellType.AllowSpace);
            result.Add(nameof(cellType.AutoConvert), cellType.AutoConvert);
            result.Add(nameof(cellType.EditMode), cellType.EditMode);
            result.Add(nameof(cellType.Ellipsis), cellType.Ellipsis);
            result.Add(nameof(cellType.EllipsisString), cellType.EllipsisString);
            result.Add(nameof(cellType.ExcelExportFormat), cellType.ExcelExportFormat);
            result.Add(nameof(cellType.ExitOnLastChar), cellType.ExitOnLastChar);
            result.Add(nameof(cellType.FocusPosition), cellType.FocusPosition);
            result.Add(nameof(cellType.FormatString), cellType.FormatString);
            result.Add(nameof(cellType.LineSpace), cellType.LineSpace);
            result.Add(nameof(cellType.MaxLength), cellType.MaxLength);
            result.Add(nameof(cellType.MaxLengthCodePage), cellType.MaxLengthCodePage);
            result.Add(nameof(cellType.MaxLengthUnit), cellType.MaxLengthUnit);
            result.Add(nameof(cellType.MaxLineCount), cellType.MaxLineCount);
            result.Add(nameof(cellType.Multiline), cellType.Multiline);
            result.Add(nameof(cellType.PasswordChar), cellType.PasswordChar);
            result.Add(nameof(cellType.PasswordRevelationMode), cellType.PasswordRevelationMode);
            result.Add(nameof(cellType.ReadOnly), cellType.ReadOnly);
            result.Add(nameof(cellType.RecommendedValue), cellType.RecommendedValue);
            result.Add(nameof(cellType.ShowRecommendedValue), cellType.ShowRecommendedValue);
            result.Add(nameof(cellType.Static), cellType.Static);
            result.Add(nameof(cellType.UseSystemPasswordChar), cellType.UseSystemPasswordChar);
            result.Add(nameof(cellType.WrapMode), cellType.WrapMode);

            return result;
        }

        /// <summary>
        /// GcDateTimeCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">GcDateTimeCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileGcDateTimeCellTypeProperties(GcDateTimeCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.AcceptsCrLf), cellType.AcceptsCrLf);
            result.Add(nameof(cellType.ClipContent), cellType.ClipContent);
            result.Add("DefaultActiveField", cellType.SerializationDefaultActiveFieldIndex);
            result.Add(nameof(cellType.EditMode), cellType.EditMode);

            var displayFieldsArray = new JArray();
            foreach (DateDisplayFieldInfo field in cellType.DisplayFields)
            {
                var jobj = JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(field));
                jobj.AddFirst(new JProperty("Type", field.GetType().Name));
                displayFieldsArray.Add(jobj);
            }
            result.Add(nameof(cellType.DisplayFields), displayFieldsArray);

            result.Add(nameof(cellType.ExcelExportFormat), cellType.ExcelExportFormat);
            result.Add(nameof(cellType.ExitOnLastChar), cellType.ExitOnLastChar);

            var fieldsArray = new JArray();
            foreach (DateFieldInfo field in cellType.Fields)
            {
                var jobj = JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(field));
                jobj.AddFirst(new JProperty("Type", field.GetType().Name));
                fieldsArray.Add(jobj);
            }
            result.Add(nameof(cellType.Fields), fieldsArray);

            result.Add(nameof(cellType.FieldsEditMode), cellType.FieldsEditMode);
            result.Add(nameof(cellType.FocusPosition), cellType.FocusPosition);
            result.Add(nameof(cellType.MaxDate), cellType.MaxDate);
            result.Add(nameof(cellType.MaxMinBehavior), cellType.MaxMinBehavior);
            result.Add(nameof(cellType.PromptChar), cellType.PromptChar);
            result.Add(nameof(cellType.ReadOnly), cellType.ReadOnly);
            result.Add(nameof(cellType.RecommendedValue), cellType.RecommendedValue);
            result.Add(nameof(cellType.ShowLiterals), cellType.ShowLiterals);
            result.Add(nameof(cellType.ShowRecommendedValue), cellType.ShowRecommendedValue);
            result.Add(nameof(cellType.Static), cellType.Static);
            result.Add(nameof(cellType.TabAction), cellType.TabAction);
            result.Add(nameof(cellType.ValidateMode), cellType.ValidateMode);
            result.Add("DropDownAllowDrop", cellType.DropDown.AllowDrop);

            return result;
        }

        /// <summary>
        /// GcNumberCellType からプロパティ情報を逆コンパイルする。
        /// </summary>
        /// <param name="cellType">GcNumberCellType オブジェクト。</param>
        /// <returns>プロパティ情報。</returns>
        private static Dictionary<string, object> DecompileGcNumberCellTypeProperties(GcNumberCellType cellType)
        {
            var result = new Dictionary<string, object>();
            result.Add(nameof(cellType.AcceptsCrLf), cellType.AcceptsCrLf);
            result.Add(nameof(cellType.AcceptsDecimal), cellType.AcceptsDecimal);
            result.Add(nameof(cellType.AllowDeleteToNull), cellType.AllowDeleteToNull);
            result.Add(nameof(cellType.ClipContent), cellType.ClipContent);

            var displayFieldsArray = new JArray();
            foreach (NumberDisplayFieldInfo field in cellType.DisplayFields)
            {
                var jobj = JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(field));
                jobj.AddFirst(new JProperty("Type", field.GetType().Name));
                displayFieldsArray.Add(jobj);
            }
            result.Add(nameof(cellType.DisplayFields), displayFieldsArray);

            result.Add(nameof(cellType.EditMode), cellType.EditMode);
            result.Add(nameof(cellType.ExcelExportFormat), cellType.ExcelExportFormat);
            result.Add(nameof(cellType.ExitOnLastChar), cellType.ExitOnLastChar);

            var fieldsArray = new JArray();
            foreach (NumberFieldInfo field in cellType.Fields)
            {
                var jobj = JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(field));
                jobj.AddFirst(new JProperty("Type", field.GetType().Name));
                fieldsArray.Add(jobj);
            }
            result.Add(nameof(cellType.Fields), fieldsArray);

            result.Add(nameof(cellType.FocusPosition), cellType.FocusPosition);
            result.Add(nameof(cellType.MaxMinBehavior), cellType.MaxMinBehavior);
            result.Add(nameof(cellType.MaxValue), cellType.MaxValue);
            result.Add(nameof(cellType.MinValue), cellType.MinValue);
            result.Add(nameof(cellType.NegativeColor), ColorTranslator.ToHtml(cellType.NegativeColor));
            result.Add(nameof(cellType.ReadOnly), cellType.ReadOnly);
            result.Add(nameof(cellType.RecommendedValue), cellType.RecommendedValue);
            result.Add(nameof(cellType.RoundPattern), cellType.RoundPattern);
            result.Add(nameof(cellType.ShowRecommendedValue), cellType.ShowRecommendedValue);
            result.Add(nameof(cellType.Static), cellType.Static);
            result.Add(nameof(cellType.UseNegativeColor), cellType.UseNegativeColor);
            result.Add(nameof(cellType.ValueSign), cellType.ValueSign);
            result.Add("DropDownAllowDrop", cellType.DropDown.AllowDrop);

            return result;
        }

        /// <summary>
        /// Enum の安全な型変換を提供する。
        /// </summary>
        private static class EnumExt
        {
            /// <summary>
            /// Enum の安全な型変換を検証する。
            /// </summary>
            /// <typeparam name="TEnum">Enum 型。</typeparam>
            /// <param name="value">Enum に変換する文字列。</param>
            /// <param name="enumValue">変換された Enum 型の値。</param>
            /// <returns>true:成功, false:失敗。</returns>
            internal static bool TryParse<TEnum>(string value, out TEnum enumValue) where TEnum : struct
            {
                return Enum.TryParse(value, out enumValue) && Enum.IsDefined(typeof(TEnum), enumValue);
            }
        }
    }
}
