namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// 列の移動を行った情報を提供します。
    /// </summary>
    public class ColumnMoveResult
    {
        /// <summary>
        /// 新しいインスタンスを生成します。
        /// </summary>
        /// <param name="originalColumnIndex">オリジナルの列インデックス。</param>
        /// <param name="beforeColumnIndex">移動前列インデックス。</param>
        /// <param name="afterColumnIndex">移動後列インデックス。</param>
        /// <param name="dataField">列の DataField プロパティ値。</param>
        public ColumnMoveResult(int originalColumnIndex, int beforeColumnIndex, int afterColumnIndex, string dataField)
        {
            OriginalColumnIndex = originalColumnIndex;
            BeforeColumnIndex = beforeColumnIndex;
            AfterColumnIndex = afterColumnIndex;
            DataField = dataField;
        }

        /// <summary>
        /// オリジナルの列インデックスを取得します。
        /// MoveColumnKeepDataField() の実行結果では、かならず BeforeColumnIndex と同一になります。
        /// </summary>
        public int OriginalColumnIndex { get; }

        /// <summary>
        /// 移動前の列インデックスを取得します。
        /// </summary>
        public int BeforeColumnIndex { get; }

        /// <summary>
        /// 移動後の列インデックスを取得します。
        /// </summary>
        public int AfterColumnIndex { get; }

        /// <summary>
        /// 列の DataField プロパティ値を取得します。
        /// </summary>
        public string DataField { get; }
    }
}
