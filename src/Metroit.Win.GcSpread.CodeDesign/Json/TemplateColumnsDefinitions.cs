using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD Column 定義のテンプレートレイアウトを構成するルート要素を提供します。
    /// </summary>
    [JsonObject]
    public class TemplateColumnsDefinitions
    {
        /// <summary>
        /// 列情報を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public TemplateColumnDefinitions[] Columns { get; set; }
    }
}
