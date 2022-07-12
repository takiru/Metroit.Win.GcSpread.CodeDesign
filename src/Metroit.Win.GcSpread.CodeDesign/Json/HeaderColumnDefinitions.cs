using FarPoint.Win.Spread;
using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列ヘッダーの列要素を提供します。
    /// </summary>
    [JsonObject("Column")]
    public class HeaderColumnDefinitions
    {
        /// <summary>
        /// 列タイトルを取得または設定します。
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
