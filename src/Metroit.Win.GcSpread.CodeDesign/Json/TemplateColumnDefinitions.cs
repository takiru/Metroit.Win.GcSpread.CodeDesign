using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD Column 定義のテンプレートレイアウトを構成する列要素を提供します。
    /// </summary>
    [JsonObject]
    public class TemplateColumnDefinitions : ColumnDefinitions
    {
        /// <summary>
        /// テンプレート名を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public new string TemplateName { get; set; }
    }
}
