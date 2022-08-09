using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// TextBoxAlternateTextInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class TextBoxAlternateTextInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="textBoxAlternateTextInfo">TextBoxAlternateTextInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(TextBoxAlternateTextInfo textBoxAlternateTextInfo)
        {
            if (textBoxAlternateTextInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(TextBoxAlternateTextInfo.DisplayNull), AlternateTextInfoConverter.Serialize(textBoxAlternateTextInfo.DisplayNull)));
            jObj.Add(new JProperty(nameof(TextBoxAlternateTextInfo.Null), AlternateTextInfoConverter.Serialize(textBoxAlternateTextInfo.Null)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="textBoxAlternateTextInfo">TextBoxAlternateTextInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(TextBoxAlternateTextInfo textBoxAlternateTextInfo, JToken prop)
        {
            var displayNullProp = prop.SelectToken(nameof(TextBoxAlternateTextInfo.DisplayNull));
            if (displayNullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(displayNullProp, textBoxAlternateTextInfo.DisplayNull);
                var foreColor = displayNullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    textBoxAlternateTextInfo.DisplayNull.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = displayNullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    textBoxAlternateTextInfo.DisplayNull.Text = alternateTextInfo.Text;
                }
            }

            var nullProp = prop.SelectToken(nameof(TextBoxAlternateTextInfo.Null));
            if (nullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(nullProp, textBoxAlternateTextInfo.Null);
                var foreColor = nullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    textBoxAlternateTextInfo.Null.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = nullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    textBoxAlternateTextInfo.Null.Text = alternateTextInfo.Text;
                }
            }
        }
    }
}
