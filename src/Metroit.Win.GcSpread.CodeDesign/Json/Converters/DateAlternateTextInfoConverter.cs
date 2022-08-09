using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DateAlternateTextInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DateAlternateTextInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dateAlternateTextInfo">DateAlternateTextInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DateAlternateTextInfo dateAlternateTextInfo)
        {
            if (dateAlternateTextInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DateAlternateTextInfo.DisplayEmptyEra), AlternateTextInfoConverter.Serialize(dateAlternateTextInfo.DisplayEmptyEra)));
            jObj.Add(new JProperty(nameof(DateAlternateTextInfo.DisplayNull), AlternateTextInfoConverter.Serialize(dateAlternateTextInfo.DisplayNull)));
            jObj.Add(new JProperty(nameof(DateAlternateTextInfo.EmptyEra), AlternateTextInfoConverter.Serialize(dateAlternateTextInfo.EmptyEra)));
            jObj.Add(new JProperty(nameof(DateAlternateTextInfo.Null), AlternateTextInfoConverter.Serialize(dateAlternateTextInfo.Null)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="dateAlternateTextInfo">DateAlternateTextInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(DateAlternateTextInfo dateAlternateTextInfo, JToken prop)
        {
            var displayEmptyEra = prop.SelectToken(nameof(DateAlternateTextInfo.DisplayEmptyEra));
            if (displayEmptyEra != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(displayEmptyEra, dateAlternateTextInfo.DisplayEmptyEra);
                var foreColor = displayEmptyEra.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    dateAlternateTextInfo.DisplayEmptyEra.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = displayEmptyEra.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    dateAlternateTextInfo.DisplayEmptyEra.Text = alternateTextInfo.Text;
                }
            }

            var displayNullProp = prop.SelectToken(nameof(DateAlternateTextInfo.DisplayNull));
            if (displayNullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(displayNullProp, dateAlternateTextInfo.DisplayNull);
                var foreColor = displayNullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    dateAlternateTextInfo.DisplayNull.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = displayNullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    dateAlternateTextInfo.DisplayNull.Text = alternateTextInfo.Text;
                }
            }

            var emptyEra = prop.SelectToken(nameof(DateAlternateTextInfo.EmptyEra));
            if (emptyEra != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(displayEmptyEra, dateAlternateTextInfo.EmptyEra);
                var foreColor = displayEmptyEra.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    dateAlternateTextInfo.EmptyEra.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = displayEmptyEra.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    dateAlternateTextInfo.EmptyEra.Text = alternateTextInfo.Text;
                }
            }

            var nullProp = prop.SelectToken(nameof(DateAlternateTextInfo.Null));
            if (nullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(nullProp, dateAlternateTextInfo.Null);
                var foreColor = nullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    dateAlternateTextInfo.Null.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = nullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    dateAlternateTextInfo.Null.Text = alternateTextInfo.Text;
                }
            }
        }
    }
}
