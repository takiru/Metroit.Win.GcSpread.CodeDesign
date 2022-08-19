using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DayOfWeekStyle オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DayOfWeekStyleConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dayOfWeekStyle">DayOfWeekStyle オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DayOfWeekStyle dayOfWeekStyle)
        {
            if (dayOfWeekStyle == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DayOfWeekStyle.DayName), dayOfWeekStyle.DayName));
            jObj.Add(new JProperty(nameof(DayOfWeekStyle.ReflectToTitle), dayOfWeekStyle.ReflectToTitle));
            jObj.Add(new JProperty(nameof(DayOfWeekStyle.SubStyle), SubStyleConverter.Serialize(dayOfWeekStyle.SubStyle)));
            jObj.Add(new JProperty(nameof(DayOfWeekStyle.WeekFlags), dayOfWeekStyle.WeekFlags));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>DayOfWeekStyle オブジェクト。</returns>
        public static DayOfWeekStyle Deserialize(JToken prop)
        {
            return Deserialize(prop, new DayOfWeekStyle());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする DayOfWeekStyle オブジェクト。</param>
        /// <returns>DayOfWeekStyle オブジェクト。</returns>
        public static DayOfWeekStyle Deserialize(JToken prop, DayOfWeekStyle source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new DayOfWeekStyle();
            if (source != null)
            {
                result.DayName = source.DayName;
                result.ReflectToTitle = source.ReflectToTitle;
                result.SubStyle = source.SubStyle;
                result.WeekFlags = source.WeekFlags;
            }

            var dayName = prop.SelectToken(nameof(DayOfWeekStyle.DayName));
            if (dayName != null)
            {
                result.DayName = dayName.ToObject<string>();
            }

            var reflectToTitle = prop.SelectToken(nameof(DayOfWeekStyle.ReflectToTitle));
            if (reflectToTitle != null)
            {
                result.ReflectToTitle = reflectToTitle.ToObject<ReflectTitle>();
            }

            var subStyle = prop.SelectToken(nameof(DayOfWeekStyle.SubStyle));
            if (subStyle != null)
            {
                result.SubStyle = SubStyleConverter.Deserialize(subStyle, result.SubStyle);
            }

            var weekFlags = prop.SelectToken(nameof(DayOfWeekStyle.WeekFlags));
            if (weekFlags != null)
            {
                result.WeekFlags = weekFlags.ToObject<WeekFlags>();
            }

            return result;
        }
    }
}
