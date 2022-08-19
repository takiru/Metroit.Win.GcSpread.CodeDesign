using FarPoint.Win.Spread;
using System.Collections.Generic;

namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// GrapeCity SPREAD SheetView の拡張メソッドを提供します。
    /// </summary>
    public static class SheetViewExtensions
    {
        /// <summary>
        /// DataField プロパティを保持したまま列の移動を行います。
        /// 列タイトルのセル結合を保持するため、SheetView.MoveColumn() の第四引数 moveContent は true として実行されます。
        /// </summary>
        /// <param name="sheetView">SheetView オブジェクト。</param>
        /// <param name="fromIndex">移動元の列インデックス。</param>
        /// <param name="toIndex">移動先の列インデックス。</param>
        /// <param name="columnCount">移動する列数。</param>
        /// <returns>列の移動を行った情報のリスト。移動が行われなかった時は空のリストを返却します。</returns>
        public static List<ColumnMoveResult> MoveColumnKeepDataField(this SheetView sheetView, int fromIndex, int toIndex, int columnCount)
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
                if (string.IsNullOrEmpty(columnMoveResult.DataField))
                {
                    continue;
                }
                sheetView.BindDataColumn(columnMoveResult.AfterColumnIndex, columnMoveResult.DataField);
            }

            return result;
        }
    }
}
