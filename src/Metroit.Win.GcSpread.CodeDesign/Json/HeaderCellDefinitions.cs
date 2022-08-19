using FarPoint.Win.Spread;
using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列ヘッダーのセル要素を提供します。
    /// </summary>
    [JsonObject("Cell")]
    public class HeaderCellDefinitions
    {
        /// <summary>
        /// テンプレート名を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string BaseTemplate { get; set; }

        /// <summary>
        /// セルタイトルを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Value { get; set; }

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
    }
}
