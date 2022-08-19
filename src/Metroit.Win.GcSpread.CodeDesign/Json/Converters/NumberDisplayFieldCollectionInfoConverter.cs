using GrapeCity.Win.Spread.InputMan.CellType;
using GrapeCity.Win.Spread.InputMan.CellType.Fields;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// NumberDisplayFieldCollectionInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal class NumberDisplayFieldCollectionInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="numberDisplayFieldCollectionInfo">NumberDisplayFieldCollectionInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(NumberDisplayFieldCollectionInfo numberDisplayFieldCollectionInfo)
        {
            if (numberDisplayFieldCollectionInfo == null)
            {
                return null;
            }

            var result = new JArray();

            foreach (NumberDisplayFieldInfo field in numberDisplayFieldCollectionInfo)
            {
                var jobj = JObject.Parse(JsonConvert.SerializeObject(field));
                jobj.AddFirst(new JProperty("Type", field.GetType().Name));
                result.Add(jobj);
            }

            return result;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <returns>NumberDisplayFieldCollectionInfo オブジェクト。</returns>
        public static NumberDisplayFieldCollectionInfo Deserialize(JArray props)
        {
            var result = new NumberDisplayFieldCollectionInfo();

            foreach (JToken prop in props)
            {
                var typeName = prop.SelectToken("Type").ToObject<string>();
                if (string.Compare(typeName, nameof(NumberDecimalGeneralFormatDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberDecimalGeneralFormatDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberDecimalPartDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberDecimalPartDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberDecimalSeparatorDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberDecimalSeparatorDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberIntegerPartDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberIntegerPartDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberLiteralDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberLiteralDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberMoneyPatternDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberMoneyPatternDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberSignDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberSignDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberSystemFormatDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberSystemFormatDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
            }

            return result;
        }
    }
}
