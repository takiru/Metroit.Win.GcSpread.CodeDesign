using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// SubStyle オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class SubStyleConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="holidayCollection">SubStyle オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(SubStyle subStyle)
        {
            if (subStyle == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(SubStyle.BackColor), ColorTranslator.ToHtml(subStyle.BackColor)));
            jObj.Add(new JProperty(nameof(SubStyle.Bold), subStyle.Bold));
            jObj.Add(new JProperty(nameof(SubStyle.ForeColor), ColorTranslator.ToHtml(subStyle.ForeColor)));
            jObj.Add(new JProperty(nameof(SubStyle.UnderLine), subStyle.UnderLine));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>SubStyle オブジェクト。</returns>
        public static SubStyle Deserialize(JToken prop)
        {
            return Deserialize(prop, new SubStyle());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする SubStyle オブジェクト。</param>
        /// <returns>SubStyle オブジェクト。</returns>
        public static SubStyle Deserialize(JToken prop, SubStyle source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new SubStyle();
            if (source != null)
            {
                result.BackColor = source.BackColor;
                result.Bold = source.Bold;
                result.ForeColor = source.ForeColor;
                result.UnderLine = source.UnderLine;
            }

            var backColor = prop.SelectToken(nameof(SubStyle.BackColor));
            if (backColor != null)
            {
                result.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var bold = prop.SelectToken(nameof(SubStyle.Bold));
            if (bold != null)
            {
                result.Bold = bold.ToObject<bool>();
            }

            var foreColor = prop.SelectToken(nameof(SubStyle.ForeColor));
            if (foreColor != null)
            {
                result.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var underLine = prop.SelectToken(nameof(SubStyle.UnderLine));
            if (underLine != null)
            {
                result.UnderLine = underLine.ToObject<bool>();
            }

            return result;
        }
    }
}
