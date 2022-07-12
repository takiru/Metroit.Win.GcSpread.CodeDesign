using Newtonsoft.Json;

namespace Metroit.Win.GcSpread.CodeDesign.Json
{
    /// <summary>
    /// GrapeCity SPREAD SheetView の列ヘッダーの特定要素のみを構成するルート要素を提供します。
    /// </summary>
    [JsonObject]
    public class SpecialColumnHeaderDefinitions
    {
        /// <summary>
        /// 列ヘッダーの行情報を提供します。
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public SpecialHeaderRowDefinitions[] Rows { get; set; }
    }
}
