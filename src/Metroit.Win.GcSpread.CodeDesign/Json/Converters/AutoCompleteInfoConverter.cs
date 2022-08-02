using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// AutoCompleteInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class AutoCompleteInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="autoCompleteInfo">AutoCompleteInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(AutoCompleteInfo autoCompleteInfo)
        {
            if (autoCompleteInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(autoCompleteInfo.CandidateListItemFont), FontConverter.Serialize(autoCompleteInfo.CandidateListItemFont)));
            jObj.Add(new JProperty(nameof(autoCompleteInfo.HighlightMatchedText), autoCompleteInfo.HighlightMatchedText));
            jObj.Add(new JProperty(nameof(autoCompleteInfo.HighlightStyle), AutoCompleteHighlightStyleInfoConverter.Serialize(autoCompleteInfo.HighlightStyle)));
            jObj.Add(new JProperty(nameof(autoCompleteInfo.MatchingMode), autoCompleteInfo.MatchingMode));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="autoCompleteInfo">AutoCompleteInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(AutoCompleteInfo autoCompleteInfo, JToken prop)
        {
            var candidateListItemFont = prop.SelectToken(nameof(AutoCompleteInfo.CandidateListItemFont));
            if (candidateListItemFont != null)
            {
                var fontObj = FontConverter.Deserialize(candidateListItemFont);
                if (fontObj != null)
                {
                    autoCompleteInfo.CandidateListItemFont = fontObj;
                }
            }

            var highlightMatchedText = prop.SelectToken(nameof(AutoCompleteInfo.HighlightMatchedText));
            if (highlightMatchedText != null)
            {
                autoCompleteInfo.HighlightMatchedText = highlightMatchedText.ToObject<bool>();
            }

            var highlightStyle = prop.SelectToken(nameof(AutoCompleteInfo.HighlightStyle));
            if (highlightStyle != null)
            {
                var highlightStyleObj = AutoCompleteHighlightStyleInfoConverter.Deserialize(highlightStyle, autoCompleteInfo.HighlightStyle);
                if (highlightStyle.SelectToken(nameof(AutoCompleteInfo.HighlightStyle.BackColor)) != null)
                {
                    autoCompleteInfo.HighlightStyle.BackColor = highlightStyleObj.BackColor;
                }
                if (highlightStyle.SelectToken(nameof(AutoCompleteInfo.HighlightStyle.Font)) != null)
                {
                    autoCompleteInfo.HighlightStyle.Font = highlightStyleObj.Font;
                }
                if (highlightStyle.SelectToken(nameof(AutoCompleteInfo.HighlightStyle.ForeColor)) != null)
                {
                    autoCompleteInfo.HighlightStyle.ForeColor = highlightStyleObj.ForeColor;
                }
            }

            var matchingMode = prop.SelectToken(nameof(AutoCompleteInfo.MatchingMode));
            if (matchingMode != null)
            {
                autoCompleteInfo.MatchingMode = matchingMode.ToObject<AutoCompleteMatchingMode>();
            }
        }
    }
}
