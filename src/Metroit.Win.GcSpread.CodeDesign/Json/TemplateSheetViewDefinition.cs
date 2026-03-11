using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のテンプレートレイアウトを構成するルート要素を提供します。
    /// </summary>
    [JsonObject]
    public class TemplateSheetViewDefinition
    {
        /// <summary>
        /// 列ヘッダー情報を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public TemplateColumnHeaderDefinition ColumnHeader { get; set; }

        /// <summary>
        /// 列情報を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public TemplateColumnDefinition[] Columns { get; set; }
    }
}
