using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ComboBoxAlternateTextInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ComboBoxAlternateTextInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="comboBoxAlternateTextInfo">ComboBoxAlternateTextInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(ComboBoxAlternateTextInfo comboBoxAlternateTextInfo)
        {
            if (comboBoxAlternateTextInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(ComboBoxAlternateTextInfo.DisplayNull), AlternateTextInfoConverter.Serialize(comboBoxAlternateTextInfo.DisplayNull)));
            jObj.Add(new JProperty(nameof(ComboBoxAlternateTextInfo.Null), AlternateTextInfoConverter.Serialize(comboBoxAlternateTextInfo.Null)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="comboBoxAlternateTextInfo">ComboBoxAlternateTextInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(ComboBoxAlternateTextInfo comboBoxAlternateTextInfo, JToken prop)
        {
            var displayNullProp = prop.SelectToken(nameof(ComboBoxAlternateTextInfo.DisplayNull));
            if (displayNullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(displayNullProp, comboBoxAlternateTextInfo.DisplayNull);
                var foreColor = displayNullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    comboBoxAlternateTextInfo.DisplayNull.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = displayNullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    comboBoxAlternateTextInfo.DisplayNull.Text = alternateTextInfo.Text;
                }
            }

            var nullProp = prop.SelectToken(nameof(ComboBoxAlternateTextInfo.Null));
            if (nullProp != null)
            {
                var alternateTextInfo = AlternateTextInfoConverter.Deserialize(nullProp, comboBoxAlternateTextInfo.Null);
                var foreColor = nullProp.SelectToken(nameof(AlternateTextInfo.ForeColor));
                if (foreColor != null)
                {
                    comboBoxAlternateTextInfo.Null.ForeColor = alternateTextInfo.ForeColor;
                }
                var text = nullProp.SelectToken(nameof(AlternateTextInfo.Text));
                if (text != null)
                {
                    comboBoxAlternateTextInfo.Null.Text = alternateTextInfo.Text;
                }
            }
        }
    }
}
