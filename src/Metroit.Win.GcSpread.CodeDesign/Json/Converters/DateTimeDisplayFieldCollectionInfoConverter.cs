using GrapeCity.Win.Spread.InputMan.CellType;
using GrapeCity.Win.Spread.InputMan.CellType.Fields;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DateTimeDisplayFieldCollectionInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal class DateTimeDisplayFieldCollectionInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dateTimeDisplayFieldCollectionInfo">DateTimeDisplayFieldCollectionInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(DateTimeDisplayFieldCollectionInfo dateTimeDisplayFieldCollectionInfo)
        {
            if (dateTimeDisplayFieldCollectionInfo == null)
            {
                return null;
            }

            var result = new JArray();

            foreach (NumberDisplayFieldInfo field in dateTimeDisplayFieldCollectionInfo)
            {
                var jobj = JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(field));
                jobj.AddFirst(new JProperty("Type", field.GetType().Name));
                result.Add(jobj);
            }

            return result;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>NumberDisplayFieldCollectionInfo オブジェクト。</returns>
        public static DateTimeDisplayFieldCollectionInfo Deserialize(JArray props)
        {
            var result = new DateTimeDisplayFieldCollectionInfo();

            foreach (JToken prop in props)
            {
                var typeName = prop.SelectToken("Type").ToObject<string>();
                if (string.Compare(typeName, nameof(DateADDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateADDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateAmPmDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateAmPmDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateDayDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateDayDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateEraDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateEraDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateEraYearDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateEraYearDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateHourDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateHourDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateLiteralDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateLiteralDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateMinuteDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateMinuteDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateMonthDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateMonthDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateSecondDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateSecondDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateShortHourDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateShortHourDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateSystemFormatDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateSystemFormatDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateWeekdayDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateWeekdayDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
                if (string.Compare(typeName, nameof(DateYearDisplayFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<DateYearDisplayFieldInfo>();
                    result.Add(info);
                    continue;
                }
            }

            return result;
        }
    }
}
