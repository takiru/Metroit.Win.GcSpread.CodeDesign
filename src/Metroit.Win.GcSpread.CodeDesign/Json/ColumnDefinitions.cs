using Newtonsoft.Json;
using System.Collections.Generic;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView のレイアウトを構成する列要素を提供します。
    /// </summary>
    [JsonObject]
    public class ColumnDefinitions : ColumnDefinitionsBase
    {
        /// <summary>
        /// 利用するテンプレート名を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string BaseTemplate { get; set; }

        /// <summary>
        /// データバインドするフィールド名を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string DataField { get; set; }

        /// <summary>
        /// セルタイプを取得または設定します。
        /// </summary>
        // TODO: ColumnsSetupProvider::CreateColumnsDefinitions()で任意のプロパティを生成したいので強制をやめてよいか？
        //[JsonProperty(Required = Required.Always)]
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string CellType { get; set; }

        /// <summary>
        /// セルタイプのプロパティを取得または設定します。
        /// サポートするプロパティはセルタイプによって異なります。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public object CellTypeProps { get; set; }

        /// <summary>
        /// 外部オプション情報を取得または設定します。
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public Dictionary<string, object> Options { get; set; }
    }
}
