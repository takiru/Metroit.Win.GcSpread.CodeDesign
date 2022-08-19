using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// YearMonthFormat オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class YearMonthFormatConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="yearMonthFormat">YearMonthFormat オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(YearMonthFormat yearMonthFormat)
        {
            if (yearMonthFormat == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(YearMonthFormat.MonthFormat), yearMonthFormat.MonthFormat));
            jObj.Add(new JProperty(nameof(YearMonthFormat.YearFormat), yearMonthFormat.YearFormat));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>YearMonthFormat オブジェクト。</returns>
        public static YearMonthFormat Deserialize(JToken prop)
        {
            return Deserialize(prop, new YearMonthFormat());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする YearMonthFormat オブジェクト。</param>
        /// <returns>YearMonthFormat オブジェクト。</returns>
        public static YearMonthFormat Deserialize(JToken prop, YearMonthFormat source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new YearMonthFormat();
            if (source != null)
            {
                result.MonthFormat = source.MonthFormat;
                result.YearFormat = source.YearFormat;
            }

            var monthFormat = prop.SelectToken(nameof(YearMonthFormat.MonthFormat));
            if (monthFormat != null)
            {
                result.MonthFormat = monthFormat.ToObject<string>();
            }

            var yearFormat = prop.SelectToken(nameof(YearMonthFormat.YearFormat));
            if (yearFormat != null)
            {
                result.YearFormat = yearFormat.ToObject<string>();
            }

            return result;
        }
    }
}
