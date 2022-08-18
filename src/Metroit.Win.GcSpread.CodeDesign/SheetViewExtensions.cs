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
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayoutFrom(this SheetView sheetView, string layoutPath,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            BindJsonLayoutFrom(sheetView, layoutPath, string.Empty, columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONファイルでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPath">テンプレートJSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayoutFrom(this SheetView sheetView, string layoutPath,
            string templateLayoutPath,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            BindJsonLayoutFrom(sheetView, layoutPath, new[] { templateLayoutPath }, columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONファイルでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templatePaths">テンプレートJSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayoutFrom(this SheetView sheetView, string layoutPath,
            string[] templateLayoutPaths,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
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
            BindJsonLayout(sheetView, layout, templateLayouts.ToArray(), columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONテキストでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="layout">バインドするJSONテキスト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayout(this SheetView sheetView, string layout,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            BindJsonLayout(sheetView, layout, string.Empty, columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONテキストでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="layout">バインドするJSONテキスト。</param>
        /// <param name="templateLayout">テンプレートJSONテキスト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayout(this SheetView sheetView, string layout,
            string templateLayout,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            BindJsonLayout(sheetView, layout, new[] { templateLayout }, columnHeaderCellTag, interceptor);
        }

        /// <summary>
        /// 指定したJSONテキストでシートをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="layout">バインドするJSONテキスト。</param>
        /// <param name="templateLayouts">テンプレートJSONテキスト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayout(this SheetView sheetView, string layout,
            string[] templateLayouts,
            Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            var layoutDefs = Newtonsoft.Json.JsonConvert.DeserializeObject<SheetViewDefinitions>(layout);
            var templateLayoutDefsList = DeserializeTemplateLayouts(templateLayouts);

            // バインド準備
            PrepareBinding(sheetView);

            // ヘッダーを編集
            var headerProvider = new ColumnHeaderSetupProvider(sheetView);
            headerProvider.SetupRows(layoutDefs.ColumnHeader.Rows, templateLayoutDefsList);
            headerProvider.SetupCells(layoutDefs.ColumnHeader.Cells, templateLayoutDefsList, columnHeaderCellTag);

            // 列を編集
            var columnProvider = new ColumnsSetupProvider(sheetView);
            columnProvider.SetupColumns(layoutDefs.Columns, layoutDefs.AllColumn, templateLayoutDefsList);
            //CompileColumns(sheetView, layoutDefs, templateLayoutDefsList);

            // JSONデータを超えて列移動やらタイトル変更やらを行いたい場合のインターセプター
            var bindResults = interceptor?.Invoke(sheetView, layoutDefs);

            // ヘッダースパンを編集
            headerProvider.SetupSpans(layoutDefs.ColumnHeader.Spans, bindResults);
        }

        ///// <summary>
        ///// 指定したJSONファイルでシートのヘッダー部のみをバインドします。
        ///// </summary>
        ///// <param name="sheetView">SheetView オブジェクト。</param>
        ///// <param name="path">JSONファイルパス。</param>
        ///// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        ///// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        ///// <exception cref="InvalidOperationException"></exception>
        //public static void BindJsonLayoutWithoutColumns(this SheetView sheetView, string layoutPath,
        //    Func<Cell, object> columnHeaderCellTag = null,
        //    Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        //{
        //    var jsonObj = JsonConvert.DeserializeSheetView(File.ReadAllText(path));
        //    BindJsonLayoutWithoutColumns(sheetView, jsonObj, columnHeaderCellTag, interceptor);
        //}

        /// <summary>
        /// 指定したJSONオブジェクトでシートのヘッダー部のみをバインドします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="jsonObj">バインドするJSONオブジェクト。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void BindJsonLayoutWithoutColumns(this SheetView sheetView, string layout,
            string[] templateLayouts, Func<Cell, object> columnHeaderCellTag = null,
            Func<SheetView, SheetViewDefinitions, IList<ColumnMoveResult>> interceptor = null)
        {
            var layoutDefs = Newtonsoft.Json.JsonConvert.DeserializeObject<SheetViewDefinitions>(layout);
            var templateLayoutDefsList = DeserializeTemplateLayouts(templateLayouts);

            // バインド準備
            PrepareBinding(sheetView);

            // ヘッダーを編集
            var header = new ColumnHeaderSetupProvider(sheetView);
            header.SetupRows(layoutDefs.ColumnHeader.Rows);
            header.SetupCells(layoutDefs.ColumnHeader.Cells, columnHeaderCellTag);

            // JSONデータを超えて列移動やらタイトル変更やらを行いたい場合のインターセプター
            var bindResults = interceptor?.Invoke(sheetView, layoutDefs);

            // ヘッダースパンを編集
            header.SetupSpans(layoutDefs.ColumnHeader.Spans, bindResults);
        }

        private static List<TemplateSheetViewDefinitions> DeserializeTemplateLayouts(string[] templateLayouts)
        {
            var templateLayoutDefsList = new List<TemplateSheetViewDefinitions>();
            if (templateLayouts.Length > 0)
            {
                foreach (var templateLayout in templateLayouts)
                {
                    if (!string.IsNullOrEmpty(templateLayout))
                    {
                        templateLayoutDefsList.Add(Newtonsoft.Json.JsonConvert.DeserializeObject<TemplateSheetViewDefinitions>(templateLayout));
                    }
                }
            }

            return templateLayoutDefsList;
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
        /// シート情報からシート定義オブジェクトを生成します。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="columnOptions">列のオプション情報設定アクションデリゲート。</param>
        /// <param name="includeHeaderRowProps">ヘッダー行情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="includeHeaderCellProps">ヘッダーセル情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="includeColumnProps">列情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="outputCellTypeProps">セルタイププロパティを生成するかどうか。</param>
        /// <param name="outputOptions">オプション情報を生成するかどうか。false の場合、第一引数は動作しません。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public static SheetViewDefinitions CreateSheetDefinitions(this SheetView sheetView, Func<Column, Dictionary<string, object>> columnOptions = null,
            string[] includeHeaderRowProps = null, string[] includeHeaderCellProps = null,
            string[] includeColumnProps = null, bool outputCellTypeProps = true, bool outputOptions = true)
        {
            var result = new SheetViewDefinitions()
            {
                ColumnHeader = new ColumnHeaderDefinitions()
            };

            var header = new ColumnHeaderSetupProvider(sheetView);

            // ヘッダー行情報
            result.ColumnHeader.Rows = header.CreateRowDefinitions(includeHeaderRowProps);

            // ヘッダーセル情報
            result.ColumnHeader.Cells = header.CreateCellDefinitions(includeHeaderCellProps);

            // ヘッダースパン情報
            result.ColumnHeader.Spans = header.CreateSpanDefinitions();

            // 列情報
            var columnProvider = new ColumnsSetupProvider(sheetView);
            result.Columns = columnProvider.CreateColumnsDefinitions(columnOptions, includeColumnProps, outputCellTypeProps, outputOptions);

            return result;
        }

        /// <summary>
        /// シート情報からシリアライズします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="columnOptions">列のオプション情報設定アクションデリゲート。</param>
        /// <param name="includeHeaderRowProps">ヘッダー行情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="includeHeaderCellProps">ヘッダーセル情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="includeColumnProps">列情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="outputCellTypeProps">セルタイププロパティを生成するかどうか。</param>
        /// <param name="outputOptions">オプション情報を生成するかどうか。false の場合、第一引数は動作しません。</param>
        /// <returns>JSONデータ文字列。</returns>
        public static string SerializeJson(this SheetView sheetView, Func<Column, Dictionary<string, object>> columnOptions = null,
            string[] includeHeaderRowProps = null, string[] includeHeaderCellProps = null,
            string[] includeColumnProps = null, bool outputCellTypeProps = true, bool outputOptions = true)
        {
            var defs = CreateSheetDefinitions(sheetView, columnOptions, includeHeaderRowProps, includeHeaderCellProps, includeColumnProps, outputCellTypeProps, outputOptions);
            return Newtonsoft.Json.JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// シート定義オブジェクトからシリアライズします。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="defs">シートの定義情報オブジェクト。</param>
        /// <returns>JSONデータ文字列。</returns>
        public static string SerializeJson(this SheetView sheetView, SheetViewDefinitions defs)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(defs);
        }

        /// <summary>
        /// シート情報から列ヘッダー行定義オブジェクトを生成します。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="includeProps">ヘッダー行情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public static HeaderRowDefinitions[] CreateHeaderRowDefinitions(this SheetView sheetView,
            string[] includeProps = null)
        {
            var header = new ColumnHeaderSetupProvider(sheetView);
            return header.CreateRowDefinitions(includeProps);
        }

        /// <summary>
        /// シート情報から列ヘッダーセル定義オブジェクトを生成します。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="includeProps">ヘッダーセル情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public static HeaderCellDefinitions[][] CreateHeaderCellDefinitions(this SheetView sheetView,
            string[] includeProps = null)
        {
            var header = new ColumnHeaderSetupProvider(sheetView);
            return header.CreateCellDefinitions(includeProps);
        }

        /// <summary>
        /// シート情報から列定義オブジェクトを生成します。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="includeProps">列情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public static ColumnDefinitions[] CreateColumnsDefinitions(this SheetView sheetView,
            Func<Column, Dictionary<string, object>> columnOptions,
            string[] includeProps = null, bool outputCellTypeProps = true, bool outputOptions = true)
        {
            var columns = new ColumnsSetupProvider(sheetView);
            return columns.CreateColumnsDefinitions(columnOptions, includeProps, outputCellTypeProps, outputOptions);
        }

        ///// <summary>
        ///// シート情報から列ヘッダーの特定のみのJSONデータオブジェクトを生成します。
        ///// </summary>
        ///// <param name="sheetView">SheetView オブジェクト。</param>
        ///// <returns>JSONデータオブジェクト。</returns>
        //public static SpecialColumnHeaderDefinitions CreateSpecialColumnHeaderDefinitions(this SheetView sheetView)
        //{
        //    var result = new SpecialColumnHeaderDefinitions();

        //    var headerRows = new List<SpecialHeaderRowDefinitions>();
        //    foreach (Row row in sheetView.ColumnHeader.Rows)
        //    {
        //        var headerRow = new SpecialHeaderRowDefinitions();

        //        var headerColumns = new List<SpecialHeaderColumnDefinitions>();
        //        foreach (Column column in sheetView.ColumnHeader.Columns)
        //        {
        //            headerColumns.Add(new SpecialHeaderColumnDefinitions()
        //            {
        //                Value = sheetView.ColumnHeader.Cells[row.Index, column.Index].Value?.ToString()
        //            });
        //        }
        //        headerRow.Columns = headerColumns.ToArray();
        //        headerRows.Add(headerRow);
        //    }
        //    result.Rows = headerRows.ToArray();

        //    return result;
        //}

        ///// <summary>
        ///// シート情報から列の特定のみのJSONデータオブジェクトを生成します。
        ///// </summary>
        ///// <param name="sheetView">SheetView オブジェクト。</param>
        ///// <returns>JSONデータオブジェクト。</returns>
        //public static SpecialColumnsDefinitions CreateSpecialColumnDefinitions(this SheetView sheetView)
        //{
        //    var result = new SpecialColumnsDefinitions();

        //    var columns = new List<SpecialColumnDefinitions>();

        //    foreach (Column svColumn in sheetView.Columns)
        //    {
        //        var column = new SpecialColumnDefinitions()
        //        {
        //            Width = svColumn.Width,
        //            Visible = svColumn.Visible,
        //            DataField = svColumn.DataField
        //        };
        //        columns.Add(column);
        //    }
        //    result.Columns = columns.ToArray();

        //    return result;
        //}

        /// <summary>
        /// バインド準備を行う。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        private static void PrepareBinding(SheetView sheetView)
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
    }
}
