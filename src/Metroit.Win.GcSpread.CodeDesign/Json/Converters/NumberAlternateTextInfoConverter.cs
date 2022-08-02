using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// NumberAlternateTextInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class NumberAlternateTextInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="numberAlternateTextInfo">NumberAlternateTextInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(NumberAlternateTextInfo numberAlternateTextInfo)
        {
            if (numberAlternateTextInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(NumberAlternateTextInfo.DisplayNull), AlternateTextInfoConverter.Serialize(numberAlternateTextInfo.DisplayNull)));
            jObj.Add(new JProperty(nameof(NumberAlternateTextInfo.DisplayZero), AlternateTextInfoConverter.Serialize(numberAlternateTextInfo.DisplayZero)));
            jObj.Add(new JProperty(nameof(NumberAlternateTextInfo.Null), AlternateTextInfoConverter.Serialize(numberAlternateTextInfo.Null)));
            jObj.Add(new JProperty(nameof(NumberAlternateTextInfo.Zero), AlternateTextInfoConverter.Serialize(numberAlternateTextInfo.Zero)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="numberAlternateTextInfo">NumberAlternateTextInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(NumberAlternateTextInfo numberAlternateTextInfo, JToken prop)
        {
            var displayNullProp = prop.SelectToken(nameof(NumberAlternateTextInfo.DisplayNull));
            if (displayNullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(displayNullProp, numberAlternateTextInfo.DisplayNull);
                var foreColor = displayNullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    numberAlternateTextInfo.DisplayNull.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = displayNullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    numberAlternateTextInfo.DisplayNull.Text = alternateTextInfo.Text;
                }
            }

            var displayZeroProp = prop.SelectToken(nameof(NumberAlternateTextInfo.DisplayZero));
            if (displayZeroProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(displayZeroProp, numberAlternateTextInfo.DisplayZero);
                var foreColor = displayZeroProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    numberAlternateTextInfo.DisplayZero.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = displayZeroProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    numberAlternateTextInfo.DisplayZero.Text = alternateTextInfo.Text;
                }
            }

            var nullProp = prop.SelectToken(nameof(NumberAlternateTextInfo.Null));
            if (nullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(nullProp, numberAlternateTextInfo.Null);
                var foreColor = nullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    numberAlternateTextInfo.Null.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = nullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    numberAlternateTextInfo.Null.Text = alternateTextInfo.Text;
                }
            }

            var zeroProp = prop.SelectToken(nameof(NumberAlternateTextInfo.Zero));
            if (zeroProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(zeroProp, numberAlternateTextInfo.Zero);
                var foreColor = zeroProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    numberAlternateTextInfo.Zero.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = zeroProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    numberAlternateTextInfo.Zero.Text = alternateTextInfo.Text;
                }
            }
        }
    }
}
