using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView の特定要素のみを構成する列要素を提供します。
    /// </summary>
    [JsonObject]
    public class SpecialColumnDefinitions
    {
        /// <summary>
        /// データバインドするフィールド名を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string DataField { get; set; }

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
    }
}
