using FarPoint.Win.Spread;
using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列要素を提供します。
    /// </summary>
    [JsonObject]
    public abstract class ColumnDefinitionsBase
    {
        /// <summary>
        /// 水平位置を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public CellHorizontalAlignment? HorizontalAlignment { get; set; }

        /// <summary>
        /// 垂直位置を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public CellVerticalAlignment? VerticalAlignment { get; set; }

        /// <summary>
        /// 自動フィルター機能を利用するかどうかを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? AllowAutoFilter { get; set; }

        /// <summary>
        /// 自動ソート機能を利用するかどうかを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? AllowAutoSort { get; set; }

        /// <summary>
        /// 幅を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public float? Width { get; set; }

        /// <summary>
        /// 表示状態を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? Visible { get; set; }

        /// <summary>
        /// IMEモードを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public System.Windows.Forms.ImeMode? ImeMode { get; set; }

        /// <summary>
        /// ロック状態を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public bool? Locked { get; set; }
    }
}
