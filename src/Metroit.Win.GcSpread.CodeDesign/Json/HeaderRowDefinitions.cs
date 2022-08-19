using FarPoint.Win.Spread;
using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列ヘッダーの行要素を提供します。
    /// </summary>
    [JsonObject("Row")]
    public class HeaderRowDefinitions
    {
        /// <summary>
        /// テンプレート名を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string BaseTemplate { get; set; }

        /// <summary>
        /// 高さを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public float? Height { get; set; }

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
