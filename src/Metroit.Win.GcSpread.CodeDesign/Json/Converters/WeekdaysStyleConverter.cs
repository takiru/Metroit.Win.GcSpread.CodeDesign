using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// WeekdaysStyle オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class WeekdaysStyleConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="weekdaysStyle">WeekdaysStyle オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(WeekdaysStyle weekdaysStyle)
        {
            if (weekdaysStyle == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(WeekdaysStyle.Sunday), DayOfWeekStyleConverter.Serialize(weekdaysStyle.Sunday)));
            jObj.Add(new JProperty(nameof(WeekdaysStyle.Monday), DayOfWeekStyleConverter.Serialize(weekdaysStyle.Monday)));
            jObj.Add(new JProperty(nameof(WeekdaysStyle.Tuesday), DayOfWeekStyleConverter.Serialize(weekdaysStyle.Tuesday)));
            jObj.Add(new JProperty(nameof(WeekdaysStyle.Wednesday), DayOfWeekStyleConverter.Serialize(weekdaysStyle.Wednesday)));
            jObj.Add(new JProperty(nameof(WeekdaysStyle.Thursday), DayOfWeekStyleConverter.Serialize(weekdaysStyle.Thursday)));
            jObj.Add(new JProperty(nameof(WeekdaysStyle.Friday), DayOfWeekStyleConverter.Serialize(weekdaysStyle.Friday)));
            jObj.Add(new JProperty(nameof(WeekdaysStyle.Saturday), DayOfWeekStyleConverter.Serialize(weekdaysStyle.Saturday)));
            
            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>WeekdaysStyle オブジェクト。</returns>
        public static WeekdaysStyle Deserialize(JToken prop)
        {
            return Deserialize(prop, new WeekdaysStyle());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする WeekdaysStyle オブジェクト。</param>
        /// <returns>WeekdaysStyle オブジェクト。</returns>
        public static WeekdaysStyle Deserialize(JToken prop, WeekdaysStyle source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new WeekdaysStyle();
            if (source != null)
            {
                result.Friday = source.Friday;
                result.Monday = source.Monday;
                result.Saturday = source.Saturday;
                result.Sunday = source.Sunday;
                result.Thursday = source.Thursday;
                result.Tuesday = source.Tuesday;
                result.Wednesday = source.Wednesday;
            }

            var friday = prop.SelectToken(nameof(WeekdaysStyle.Friday));
            if (friday != null)
            {
                result.Friday = DayOfWeekStyleConverter.Deserialize(friday, result.Friday);
            }

            var monday = prop.SelectToken(nameof(WeekdaysStyle.Monday));
            if (monday != null)
            {
                result.Monday = DayOfWeekStyleConverter.Deserialize(monday, result.Monday);
            }

            var saturday = prop.SelectToken(nameof(WeekdaysStyle.Saturday));
            if (saturday != null)
            {
                result.Saturday = DayOfWeekStyleConverter.Deserialize(saturday, result.Saturday);
            }

            var sunday = prop.SelectToken(nameof(WeekdaysStyle.Sunday));
            if (sunday != null)
            {
                result.Sunday = DayOfWeekStyleConverter.Deserialize(sunday, result.Sunday);
            }

            var thursday = prop.SelectToken(nameof(WeekdaysStyle.Thursday));
            if (thursday != null)
            {
                result.Thursday = DayOfWeekStyleConverter.Deserialize(thursday, result.Thursday);
            }

            var tuesday = prop.SelectToken(nameof(WeekdaysStyle.Tuesday));
            if (tuesday != null)
            {
                result.Tuesday = DayOfWeekStyleConverter.Deserialize(tuesday, result.Tuesday);
            }

            var wednesday = prop.SelectToken(nameof(WeekdaysStyle.Wednesday));
            if (friday != null)
            {
                result.Wednesday = DayOfWeekStyleConverter.Deserialize(wednesday, result.Wednesday);
            }

            return result;
        }
    }
}
