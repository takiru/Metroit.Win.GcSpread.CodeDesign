using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// AutoCompleteHighlightStyleInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class AutoCompleteHighlightStyleInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="autoCompleteHighlightStyleInfo">AutoCompleteHighlightStyleInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(AutoCompleteHighlightStyleInfo autoCompleteHighlightStyleInfo)
        {
            if (autoCompleteHighlightStyleInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(autoCompleteHighlightStyleInfo.BackColor), ColorTranslator.ToHtml(autoCompleteHighlightStyleInfo.BackColor)));
            jObj.Add(new JProperty(nameof(autoCompleteHighlightStyleInfo.Font), FontConverter.Serialize(autoCompleteHighlightStyleInfo.Font)));
            jObj.Add(new JProperty(nameof(autoCompleteHighlightStyleInfo.ForeColor), ColorTranslator.ToHtml(autoCompleteHighlightStyleInfo.ForeColor)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>AutoCompleteHighlightStyleInfo オブジェクト。</returns>
        public static AutoCompleteHighlightStyleInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new AutoCompleteHighlightStyleInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする AutoCompleteHighlightStyleInfo オブジェクト。</param>
        /// <returns>AutoCompleteHighlightStyleInfo オブジェクト。</returns>
        public static AutoCompleteHighlightStyleInfo Deserialize(JToken prop, AutoCompleteHighlightStyleInfo source)
        {
            var result = new AutoCompleteHighlightStyleInfo();
            if (source != null)
            {
                result.BackColor = source.BackColor;
                result.Font = source.Font;
                result.ForeColor = source.ForeColor;
            }

            var backColor = prop.SelectToken(nameof(AutoCompleteHighlightStyleInfo.BackColor));
            if (backColor != null)
            {
                result.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var font = prop.SelectToken(nameof(AutoCompleteHighlightStyleInfo.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    result.Font = fontObj;
                }
            }

            var foreColor = prop.SelectToken(nameof(AutoCompleteHighlightStyleInfo.ForeColor));
            if (foreColor != null)
            {
                result.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            return result;
        }
    }
}
