using GrapeCity.Win.Spread.InputMan.CellType;
using GrapeCity.Win.Spread.InputMan.CellType.Fields;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DateTimeFieldCollectionInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal class DateTimeFieldCollectionInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dateTimeFieldCollectionInfo">DateTimeFieldCollectionInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(DateTimeFieldCollectionInfo dateTimeFieldCollectionInfo)
        {
            if (dateTimeFieldCollectionInfo == null)
            {
                return null;
            }

            var result = new JArray();

            foreach (DateFieldInfo field in dateTimeFieldCollectionInfo)
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
        /// <returns>DateTimeFieldCollectionInfo オブジェクト。</returns>
        public static DateTimeFieldCollectionInfo Deserialize(JArray props)
        {
            var result = new DateTimeFieldCollectionInfo();

            foreach (JToken prop in props)
            {
                var typeName = prop.SelectToken("Type").ToObject<string>();
                if (string.Compare(typeName, nameof(DateAmPmFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateAmPmFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateDayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateDayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateEraFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateEraFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateEraYearFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateEraYearFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateHourFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateHourFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateLiteralFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateLiteralFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateMinuteFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateMinuteFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateMonthFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateMonthFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateSecondFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateSecondFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateShortHourFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateShortHourFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateYearFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateYearFieldInfo>();
                    result.Add(info);
                    continue;
                }
            }

            return result;
        }
    }
}
