using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のテンプレートレイアウトを構成する列ヘッダー要素を提供します。
    /// </summary>
    [JsonObject]
    public class TemplateColumnHeaderDefinition
    {
        /// <summary>
        /// 列ヘッダーの行情報を提供します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public TemplateHeaderRowDefinition[] Rows { get; set; }

        /// <summary>
        /// 列ヘッダーのセル情報を提供します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public TemplateHeaderCellDefinition[] Cells { get; set; }
    }
}
