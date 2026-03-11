using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列ヘッダー要素を提供します。
    /// </summary>
    [JsonObject]
    public class ColumnHeaderDefinition
    {
        /// <summary>
        /// 列ヘッダーの行情報を提供します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public HeaderRowDefinition[] Rows { get; set; }

        /// <summary>
        /// 列ヘッダーのセル情報を提供します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public HeaderCellDefinition[][] Cells { get; set; }

        /// <summary>
        /// セル結合情報を提供します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public SpanDefinition[] Spans { get; set; }
    }
}
