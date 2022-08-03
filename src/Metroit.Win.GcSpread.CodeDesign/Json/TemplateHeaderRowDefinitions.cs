using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列ヘッダーの行要素を提供します。
    /// </summary>
    [JsonObject("Row")]
    public class TemplateHeaderRowDefinitions: HeaderRowDefinitions
    {
        /// <summary>
        /// テンプレート名を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public string TemplateName { get; set; }
    }
}
