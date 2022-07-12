using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成するセル結合要素を提供します。
    /// </summary>
    [JsonObject("Span")]
    public class SpanDefinitions
    {
        /// <summary>
        /// 行インデックスを取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public int Row { get; set; }

        /// <summary>
        /// 列インデックスを取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public int Column { get; set; }

        /// <summary>
        /// 行数を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public int RowCount { get; set; }

        /// <summary>
        /// 列数を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public int ColumnCount { get; set; }
    }
}
