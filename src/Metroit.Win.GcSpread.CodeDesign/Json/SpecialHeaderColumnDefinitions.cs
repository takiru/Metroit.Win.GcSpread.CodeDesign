using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView の列ヘッダーの特定列要素のみを提供します。
    /// </summary>
    [JsonObject("Column")]
    public class SpecialHeaderColumnDefinitions
    {
        /// <summary>
        /// 列タイトルを取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Value { get; set; }
    }
}
