using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のテンプレートレイアウトを構成する列ヘッダーのセル要素を提供します。
    /// </summary>
    [JsonObject("Cell")]
    public class TemplateHeaderCellDefinitions : HeaderCellDefinitions
    {
        /// <summary>
        /// テンプレート名を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public string TemplateName { get; set; }
    }
}
