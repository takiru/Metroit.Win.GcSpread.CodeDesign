using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView の列の特定要素のみを構成するルート要素を提供します。
    /// </summary>
    [JsonObject]
    public class SpecialColumnsDefinitions
    {
        /// <summary>
        /// 列情報を取得または設定します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public SpecialColumnDefinitions[] Columns { get; set; }
    }
}
