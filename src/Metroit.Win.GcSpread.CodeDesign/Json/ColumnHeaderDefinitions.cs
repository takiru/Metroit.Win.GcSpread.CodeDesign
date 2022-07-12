using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列ヘッダー要素を提供します。
    /// </summary>
    [JsonObject]
    public class ColumnHeaderDefinitions
    {
        /// <summary>
        /// 列ヘッダーの行情報を提供します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public HeaderRowDefinitions[] Rows { get; set; }

        /// <summary>
        /// セル結合情報を提供します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public SpanDefinitions[] Spans { get; set; }
    }
}
