using Metroit.Win.GcSpread.CodeDesign.Json;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            var jsonObj = JsonConvert.DeserializeSheetView(File.ReadAllText(path));
            BindJsonLayout(sheetView, jsonObj, columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONオブジェクトでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="jsonObj">バインドするJSONオブジェクト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayout(this SheetView sheetView, SheetViewDefinitions jsonObj,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            // バインド準備
            PrepareCompilation(sheetView);

            // ヘッダーを編集
            ComplieColumnHeader(sheetView, jsonObj.ColumnHeader, columnHeaderCellTag);

            // 列を編集
            CompileColumns(sheetView, jsonObj);

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
        private static void CompileColumns(SheetView sheetView, SheetViewDefinitions root)
        {
            var columns = root.Columns;
            if (columns == null)
            {
                throw new ArgumentNullException("Columns");
            }

            foreach (var column in columns.Select((Item, Index) => new { Item, Index }))
            {
                ICellType cellType;
                switch (column.Item.CellType.ToLower())
                {
                    case "textcelltype":
                        cellType = new TextCellType();
                        sheetView.Columns[column.Index].CellType = cellType;
                        CompileTextCellType((TextCellType)cellType, column.Item.CellTypeProperties);
                        break;

                    case "datetimecelltype":
                        cellType = new DateTimeCellType();
                        sheetView.Columns[column.Index].CellType = cellType;
                        CompileDateTimeCellType((DateTimeCellType)cellType, column.Item.CellTypeProperties);
                        break;

                    case "numbercelltype":
                        cellType = new NumberCellType();
                        sheetView.Columns[column.Index].CellType = cellType;
                        CompileNumberCellType((NumberCellType)cellType, column.Item.CellTypeProperties);
                        break;

                    case "comboboxcelltype":
                        cellType = new ComboBoxCellType();
                        sheetView.Columns[column.Index].CellType = cellType;
                        CompileComboBoxCellType((ComboBoxCellType)cellType, column.Item.CellTypeProperties);
                        break;

                    case "checkboxcelltype":
                        cellType = new CheckBoxCellType();
                        sheetView.Columns[column.Index].CellType = cellType;
                        CompileCheckBoxCellType((CheckBoxCellType)cellType, column.Item.CellTypeProperties);
                        break;

                    case "buttoncelltype":
                        cellType = new ButtonCellType();
                        sheetView.Columns[column.Index].CellType = cellType;
                        CompileButtonCellType((ButtonCellType)cellType, column.Item.CellTypeProperties);
                        break;
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
                    column.CellTypeProperties = DecompileTextCellTypeProperties((TextCellType)svColumn.CellType);
                }
                if (svColumn.CellType is DateTimeCellType)
                {
                    column.CellTypeProperties = DecompileDateTimeCellTypeProperties((DateTimeCellType)svColumn.CellType);
                }
                if (svColumn.CellType is NumberCellType)
                {
                    column.CellTypeProperties = DecompileNumberCellTypeProperties((NumberCellType)svColumn.CellType);
                }
                if (svColumn.CellType is ComboBoxCellType)
                {
                    column.CellTypeProperties = DecompileComboBoxCellTypeProperties((ComboBoxCellType)svColumn.CellType);
                }
                if (svColumn.CellType is CheckBoxCellType)
                {
                    column.CellTypeProperties = DecompileCheckBoxCellTypeProperties((CheckBoxCellType)svColumn.CellType);
                }
                if (svColumn.CellType is ButtonCellType)
                {
                    column.CellTypeProperties = DecompileButtonCellTypeProperties((ButtonCellType)svColumn.CellType);
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
