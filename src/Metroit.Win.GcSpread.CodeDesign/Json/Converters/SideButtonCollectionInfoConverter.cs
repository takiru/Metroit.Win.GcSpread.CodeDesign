using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// SideButtonCollectionInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class SideButtonCollectionInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="sideButtonCollectionInfo">SideButtonCollectionInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(SideButtonCollectionInfo sideButtonCollectionInfo)
        {
            if (sideButtonCollectionInfo == null)
            {
                return null;
            }

            var displayFieldsArray = new JArray();

            foreach (SideButtonBaseInfo sideButtonBaseInfo in sideButtonCollectionInfo)
            {
                var jobj = JObject.Parse(JsonConvert.SerializeObject(sideButtonBaseInfo));
                jobj.AddFirst(new JProperty("Type", sideButtonBaseInfo.GetType().Name));
                displayFieldsArray.Add(jobj);
            }
            return displayFieldsArray;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <returns>SideButtonCollectionInfo オブジェクト。</returns>
        public static SideButtonCollectionInfo Deserialize(JArray props)
        {
            var result = new SideButtonCollectionInfo();

            foreach (JToken prop in props)
            {
                SideButtonBaseInfo info = null;
                var typeName = prop.SelectToken("Type").ToObject<string>();
                if (string.Compare(typeName, nameof(DropDownButtonInfo), true) == 0)
                {
                    info = prop.ToObject<DropDownButtonInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(SideButtonInfo), true) == 0)
                {
                    info = prop.ToObject<SideButtonInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(SpinButtonInfo), true) == 0)
                {
                    info = prop.ToObject<SpinButtonInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(SymbolButtonInfo), true) == 0)
                {
                    info = prop.ToObject<SymbolButtonInfo>();
                    result.Add(info);
                    continue;
                }
            }

            return result;
        }
    }
}
