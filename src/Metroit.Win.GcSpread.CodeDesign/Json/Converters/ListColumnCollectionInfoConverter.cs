using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ListColumnCollectionInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ListColumnCollectionInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="listColumnCollectionInfo">ListColumnCollectionInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(ListColumnCollectionInfo listColumnCollectionInfo)
        {
            if (listColumnCollectionInfo == null)
            {
                return null;
            }

            var result = new JArray();

            foreach (ListColumnInfo listColumnInfo in listColumnCollectionInfo)
            {
                var jobj = ListColumnInfoConverter.Serialize(listColumnInfo);
                result.Add(jobj);
            }

            return result;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="listColumnCollectionInfo">ListColumnCollectionInfo オブジェクト。</param>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <returns>ListColumnCollectionInfo オブジェクト。</returns>
        public static void Deserialize(ListColumnCollectionInfo listColumnCollectionInfo, JArray props)
        {
            if (props == null || !props.HasValues)
            {
                return;
            }

            foreach (JToken prop in props)
            {
                var listColumnInfo = ListColumnInfoConverter.Deserialize(prop);
                listColumnCollectionInfo.Add(listColumnInfo);
            }

            return;
        }
    }
}
