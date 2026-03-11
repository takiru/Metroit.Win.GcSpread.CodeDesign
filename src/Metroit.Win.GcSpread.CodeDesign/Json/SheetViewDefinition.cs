using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成するルート要素を提供します。
    /// </summary>
    [JsonObject]
    public class SheetViewDefinition
    {
        /// <summary>
        /// 列ヘッダー情報を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public ColumnHeaderDefinition ColumnHeader { get; set; }

        /// <summary>
        /// 列の全体情報を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public AllColumnDefinition AllColumn { get; set; }

        /// <summary>
        /// 列情報を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public ColumnDefinition[] Columns { get; set; }
    }
}
