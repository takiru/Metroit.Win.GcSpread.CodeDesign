using FarPoint.Win.Spread;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// SheetView の生成操作を提供します。
    /// </summary>
    public class SheetViewGenerator
    {
        /// <summary>
        /// 処理の対象となる SheetView。
        /// </summary>
        private SheetView SheetView { get; }

        /// <summary>
        /// 新しい SheetViewGenerator インスタンスを生成します。
        /// </summary>
        /// <param name="sheetView">処理対象とする SheetView。</param>
        public SheetViewGenerator(SheetView sheetView)
        {
            SheetView = sheetView;
        }

        /// <summary>
        /// シート定義オブジェクトからJSON文字列にシリアライズします。
        /// </summary>
        /// <param name="def">シート定義オブジェクト。</param>
        /// <returns>JSON文字列。</returns>
        public static string SerializeJson(SheetViewDefinition def)
        {
            return JsonConvert.SerializeObject(def);
        }

        /// <summary>
        /// JSON文字列からシート定義オブジェクトにデシリアライズします。
        /// </summary>
        /// <param name="json">シート定義JSON文字列。</param>
        /// <returns>シート定義オブジェクト。</returns>
        public static SheetViewDefinition DeserializeJson(string json)
        {
            return JsonConvert.DeserializeObject<SheetViewDefinition>(json);
        }

        /// <summary>
        /// 指定したJSONファイルでシートレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutFrom(string layoutPath,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutFrom(layoutPath, string.Empty, columnHeaderCellTag, generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSONファイルでシートレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPath">テンプレートJSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutFrom(string layoutPath, string templateLayoutPath,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutFrom(layoutPath, new[] { templateLayoutPath }, columnHeaderCellTag, generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSONファイルでシートレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPaths">テンプレートJSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutFrom(string layoutPath, string[] templateLayoutPaths,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutFrom(layoutPath, templateLayoutPaths, columnHeaderCellTag, generatedColumn, interceptor, true);
        }

        /// <summary>
        /// 指定したJSON文字列でシートレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayout(string layout,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayout(layout, string.Empty, columnHeaderCellTag, generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSON文字列でシートレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="templateLayout">テンプレートJSON文字列。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayout(string layout, string templateLayout,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayout(layout, new[] { templateLayout }, columnHeaderCellTag, generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSON文字列でシートレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="templateLayouts">テンプレートJSON文字列。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayout(string layout, string[] templateLayouts,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayout(layout, templateLayouts, columnHeaderCellTag, generatedColumn, interceptor, true);
        }

        /// <summary>
        /// 指定したJSONファイルでヘッダー部のみのシートレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutWithoutColumnsFrom(string layoutPath,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutWithoutColumnsFrom(layoutPath, string.Empty, columnHeaderCellTag,
                generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSONファイルでヘッダー部のみのシートレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPath">テンプレートJSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutWithoutColumnsFrom(string layoutPath, string templateLayoutPath,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutWithoutColumnsFrom(layoutPath, new[] { templateLayoutPath },
                columnHeaderCellTag, generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSONファイルでヘッダー部のみのシートレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPaths">テンプレートJSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutWithoutColumnsFrom(string layoutPath, string[] templateLayoutPaths,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutFrom(layoutPath, templateLayoutPaths, columnHeaderCellTag, generatedColumn,
                interceptor, false);
        }

        /// <summary>
        /// 指定したJSON文字列でヘッダー部のみのシートレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutWithoutColumns(string layout,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutWithoutColumns(layout, string.Empty, columnHeaderCellTag,
                generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSON文字列でヘッダー部のみのシートレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="templateLayout">テンプレートJSON文字列。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutWithoutColumns(string layout, string templateLayout,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayoutWithoutColumns(layout, new[] { templateLayout }, columnHeaderCellTag,
                generatedColumn, interceptor);
        }

        /// <summary>
        /// 指定したJSONオブジェクトでヘッダー部のみのシートレイアウトを生成します。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="templateLayouts">テンプレートJSON文字列。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        public void GenerateLayoutWithoutColumns(string layout, string[] templateLayouts,
            Func<Cell, object> columnHeaderCellTag = null, Action<Column> generatedColumn = null,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor = null)
        {
            GenerateLayout(layout, templateLayouts, columnHeaderCellTag, generatedColumn, interceptor, false);
        }

        /// <summary>
        /// 指定したJSONファイルでヘッダー部のみのシートレイアウトを生成します。
        /// </summary>
        /// <param name="layoutPath">JSONファイルパス。</param>
        /// <param name="templateLayoutPaths">テンプレートJSONファイルパス。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する特に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <param name="generateColumns">列定義をバインドするかどうか。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        private void GenerateLayoutFrom(string layoutPath, string[] templateLayoutPaths,
            Func<Cell, object> columnHeaderCellTag, Action<Column> generatedColumn,
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor,
            bool generateColumns)
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
            GenerateLayout(layout, templateLayouts.ToArray(), columnHeaderCellTag, generatedColumn,
                interceptor, generateColumns);
        }

        /// <summary>
        /// 指定したJSON文字列でシートレイアウトを生成する。
        /// </summary>
        /// <param name="layout">JSON文字列。</param>
        /// <param name="templateLayouts">テンプレートJSON文字列。</param>
        /// <param name="columnHeaderCellTag">列ヘッダーセルの Tag プロパティを設定する時に呼び出されます。</param>
        /// <param name="generatedColumn">列が生成されたときのデリゲート。</param>
        /// <param name="interceptor">列情報を変更するためのインターセプターデリゲート。ヘッダーのセル結合が行われる前に呼び出されます。</param>
        /// <param name="generateColumns">列定義をバインドするかどうか。</param>
        /// <exception cref="InvalidOperationException"><see cref="SheetView.AutoGenerateColumns"/>, <see cref="SheetView.DataAutoCellTypes"/>, <see cref="SheetView.DataAutoSizeColumns"/>, <see cref="SheetView.DataAutoHeadings"/> のいずれかが<see langword="true"/>です。</exception>
        private void GenerateLayout(string layout, string[] templateLayouts,
            Func<Cell, object> columnHeaderCellTag, Action<Column> generatedColumn, 
            Func<SheetView, SheetViewDefinition, IList<ColumnMoveResult>> interceptor,
            bool generateColumns)
        {
            // ジェネレート準備
            PrepareGenerating();

            var sheetDefinition = JsonConvert.DeserializeObject<SheetViewDefinition>(layout);
            var templateDefinitionList = DeserializeTemplateLayouts(templateLayouts);

            // ヘッダーをジェネレート
            var headerGenerator = new ColumnHeaderGenerator(SheetView);
            headerGenerator.GenerateRows(sheetDefinition.ColumnHeader.Rows, templateDefinitionList);
            headerGenerator.GenerateCells(sheetDefinition.ColumnHeader.Cells, templateDefinitionList, columnHeaderCellTag);

            // 列をジェネレート
            if (generateColumns)
            {
                var columnGenerator = new ColumnGenerator(SheetView);
                columnGenerator.GenerateColumns(sheetDefinition.Columns, templateDefinitionList, sheetDefinition.AllColumn, generatedColumn);
            }

            // JSONデータを超えて列移動やらタイトル変更やらを行いたい場合のインターセプター
            var bindResults = interceptor?.Invoke(SheetView, sheetDefinition);

            // ヘッダースパンをジェネレート
            headerGenerator.GenerateSpans(sheetDefinition.ColumnHeader.Spans, bindResults);
        }

        /// <summary>
        /// ジェネレート準備を行う。
        /// </summary>
        private void PrepareGenerating()
        {
            // NOTE: レイアウト読込において無効にしなければならないプロパティが有効な場合は実行不可とする
            if (SheetView.AutoGenerateColumns)
            {
                throw new InvalidOperationException("AutoGenerateColumns = true の場合、バインドを行うことはできません。");
            }
            if (SheetView.DataAutoCellTypes)
            {
                throw new InvalidOperationException("DataAutoCellTypes = true の場合、バインドを行うことはできません。");
            }
            if (SheetView.DataAutoSizeColumns)
            {
                throw new InvalidOperationException("DataAutoSizeColumns = true の場合、バインドを行うことはできません。");
            }
            if (SheetView.DataAutoHeadings)
            {
                throw new InvalidOperationException("DataAutoHeadings = true の場合、バインドを行うことはできません。");
            }

            SheetView.ColumnHeader.Columns.Count = 0;
            SheetView.ColumnHeader.Rows.Count = 0;
            SheetView.Columns.Count = 0;
        }

        /// <summary>
        /// テンプレートJSONをデシリアライズする。
        /// </summary>
        /// <param name="templateLayouts">テンプレートJSON文字列。</param>
        /// <returns>テンプレートオブジェクトのリスト。1つもテンプレートがなかったときは空のリストを返却する。</returns>
        private static List<TemplateSheetViewDefinition> DeserializeTemplateLayouts(string[] templateLayouts)
        {
            var templateDefinitionsList = new List<TemplateSheetViewDefinition>();
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
        /// シート情報からシート定義オブジェクトを生成します。
        /// </summary>
        /// <param name="columnOptions">列のオプション情報設定アクションデリゲート。</param>
        /// <param name="includeHeaderRowProps">ヘッダー行情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="includeHeaderCellProps">ヘッダーセル情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="includeColumnProps">列情報の生成に含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみが生成されます。</param>
        /// <param name="outputCellTypeProps">セルタイププロパティを生成するかどうか。</param>
        /// <param name="outputOptions">オプション情報を生成するかどうか。false の場合、第一引数は動作しません。</param>
        /// <returns>JSONデータオブジェクト。</returns>
        public SheetViewDefinition CreateSheetDefinitions(Func<Column, Dictionary<string, object>> columnOptions = null,
            string[] includeHeaderRowProps = null, string[] includeHeaderCellProps = null,
            string[] includeColumnProps = null, bool outputCellTypeProps = true, bool outputOptions = true)
        {
            var result = new SheetViewDefinition()
            {
                ColumnHeader = new ColumnHeaderDefinition()
            };

            var header = new ColumnHeaderGenerator(SheetView);

            // ヘッダー行情報
            result.ColumnHeader.Rows = header.CreateRowDefinitions(includeHeaderRowProps);

            // ヘッダーセル情報
            result.ColumnHeader.Cells = header.CreateCellDefinitions(includeHeaderCellProps);

            // ヘッダースパン情報
            result.ColumnHeader.Spans = header.CreateSpanDefinitions();

            // 列情報
            var columnProvider = new ColumnGenerator(SheetView);
            result.Columns = columnProvider.CreateColumnsDefinitions(columnOptions, includeColumnProps, outputCellTypeProps, outputOptions);

            return result;
        }
    }
}
