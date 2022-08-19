using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// HolidayCollection オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class HolidayCollectionConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="holidayCollection">HolidayCollection オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(HolidayCollection holidayCollection)
        {
            if (holidayCollection == null)
            {
                return null;
            }

            var jArray = new JArray();

            foreach (Holiday h in holidayCollection)
            {
                if (h == null)
                {
                    continue;
                }

                var jObj = new JObject();
                jObj.Add(new JProperty(nameof(Holiday.EndDay), h.EndDay));
                jObj.Add(new JProperty(nameof(Holiday.EndMonth), h.EndMonth));
                jObj.Add(new JProperty(nameof(Holiday.Name), h.Name));
                jObj.Add(new JProperty(nameof(Holiday.StartDay), h.StartDay));
                jObj.Add(new JProperty(nameof(Holiday.StartMonth), h.StartMonth));

                jArray.Add(jObj);
            }

            return jArray;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <returns>HolidayStyleCollection オブジェクト。</returns>
        public static HolidayCollection Deserialize(JArray props)
        {
            return Deserialize(props, new HolidayCollection());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする HolidayStyleCollection オブジェクト。</param>
        /// <returns>HolidayStyleCollection オブジェクト。</returns>
        public static HolidayCollection Deserialize(JArray props, HolidayCollection source)
        {
            if (props.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new HolidayCollection();
            foreach (Holiday h in source)
            {
                var newH = new Holiday(h.Name, h.StartMonth, h.StartDay, h.EndMonth, h.EndDay);
            }

            var i = 0;
            foreach (JToken prop in props)
            {
                if (result.Count - 1 < i)
                {
                    result.Add(new Holiday(1, 1));
                }

                Holiday h = result[i] as Holiday;

                var endDay = prop.SelectToken(nameof(Holiday.EndDay));
                if (endDay != null)
                {
                    h.EndDay = endDay.ToObject<int>();
                }

                var endMonth = prop.SelectToken(nameof(Holiday.EndMonth));
                if (endMonth != null)
                {
                    h.EndMonth = endMonth.ToObject<int>();
                }

                var name = prop.SelectToken(nameof(Holiday.Name));
                if (name != null)
                {
                    h.Name = name.ToObject<string>();
                }

                var startDay = prop.SelectToken(nameof(Holiday.StartDay));
                if (startDay != null)
                {
                    h.StartDay = startDay.ToObject<int>();
                }

                var startMonth = prop.SelectToken(nameof(Holiday.StartMonth));
                if (startMonth != null)
                {
                    h.StartMonth = startMonth.ToObject<int>();
                }

                i++;
            }

            return result;
        }
    }
}
