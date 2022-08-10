using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Collections;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// HolidayStyleCollection オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class HolidayStyleCollectionConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="holidayStyleCollection">HolidayStyleCollection オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(HolidayStyleCollection holidayStyleCollection)
        {
            if (holidayStyleCollection == null)
            {
                return null;
            }

            var jArray = new JArray();

            foreach (DictionaryEntry de in holidayStyleCollection)
            {
                if (de.Value == null)
                {
                    continue;
                }

                var hs = de.Value as HolidayStyle;

                var jObj = new JObject();
                jObj.Add(new JProperty(nameof(HolidayStyle.Holidays), HolidayCollectionConverter.Serialize(hs.Holidays)));
                jObj.Add(new JProperty(nameof(HolidayStyle.Name), hs.Name));
                jObj.Add(new JProperty(nameof(HolidayStyle.SubStyle), SubStyleConverter.Serialize(hs.SubStyle)));

                jArray.Add(jObj);
            }

            return jArray;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <returns>HolidayStyleCollection オブジェクト。</returns>
        public static HolidayStyleCollection Deserialize(JArray props)
        {
            return Deserialize(props, new HolidayStyleCollection());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする HolidayStyleCollection オブジェクト。</param>
        /// <returns>HolidayStyleCollection オブジェクト。</returns>
        public static HolidayStyleCollection Deserialize(JArray props, HolidayStyleCollection source)
        {
            if (props.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new HolidayStyleCollection();
            foreach (HolidayStyle hs in source)
            {
                var newHs = new HolidayStyle();
                newHs.Holidays = hs.Holidays;
                newHs.Name = hs.Name;
                newHs.SubStyle = hs.SubStyle;
            }

            var i = 0;
            foreach (JToken prop in props)
            {
                if (result.Count - 1 < i)
                {
                    result.Add(string.Empty, new HolidayStyle());
                }

                var holidays = prop.SelectToken(nameof(HolidayStyle.Holidays));
                if (holidays != null)
                {
                    result[i].Holidays = HolidayCollectionConverter.Deserialize((JArray)holidays, result[i].Holidays);
                }

                var name = prop.SelectToken(nameof(HolidayStyle.Name));
                if (name != null)
                {
                    result[i].Name = name.ToObject<string>();
                }

                var subStyle = prop.SelectToken(nameof(HolidayStyle.SubStyle));
                if (subStyle != null)
                {
                    result[i].SubStyle = SubStyleConverter.Deserialize((JArray)subStyle, result[i].SubStyle);
                }

                i++;
            }

            return result;
        }
    }
}
