using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView の列ヘッダーの特定行要素のみを提供します。
    /// </summary>
    [JsonObject("Row")]
    public class SpecialHeaderRowDefinitions
    {
        /// <summary>
        /// 列ヘッダーの列情報を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public SpecialHeaderColumnDefinitions[] Columns { get; set; }
    }
}
