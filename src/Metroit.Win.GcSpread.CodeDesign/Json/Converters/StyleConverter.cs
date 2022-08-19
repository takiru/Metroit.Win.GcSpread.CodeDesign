using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// Style オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class StyleConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="style">Style オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(Style style)
        {
            if (style == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(Style.BackColor), ColorTranslator.ToHtml(style.BackColor)));
            jObj.Add(new JProperty(nameof(Style.Bevel), JProperty.FromObject(style.Bevel)));
            jObj.Add(new JProperty(nameof(Style.BorderColor), ColorTranslator.ToHtml(style.BorderColor)));
            jObj.Add(new JProperty(nameof(Style.BorderStyle), style.BorderStyle));
            jObj.Add(new JProperty(nameof(Style.ContentAlignment), style.ContentAlignment));
            jObj.Add(new JProperty(nameof(Style.Font), FontConverter.Serialize(style.Font)));
            jObj.Add(new JProperty(nameof(Style.ForeColor), ColorTranslator.ToHtml(style.ForeColor)));
            jObj.Add(new JProperty(nameof(Style.TextEffect), style.TextEffect));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>Style オブジェクト。</returns>
        public static Style Deserialize(JToken prop)
        {
            return Deserialize(prop, new Style());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする Style オブジェクト。</param>
        /// <returns>Style オブジェクト。</returns>
        public static Style Deserialize(JToken prop, Style source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new Style();
            if (source != null)
            {
                result.BackColor = source.BackColor;
                result.Bevel = source.Bevel;
                result.BorderColor = source.BorderColor;
                result.BorderStyle = source.BorderStyle;
                result.ContentAlignment = source.ContentAlignment;
                result.Font = source.Font;
                result.ForeColor = source.ForeColor;
                result.TextEffect = source.TextEffect;
            }

            var backColor = prop.SelectToken(nameof(Style.BackColor));
            if (backColor != null)
            {
                result.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var bevel = prop.SelectToken(nameof(Style.Bevel));
            if (bevel != null)
            {
                result.Bevel = bevel.ToObject<Bevel>();
            }

            var borderColor = prop.SelectToken(nameof(Style.BorderColor));
            if (borderColor != null)
            {
                result.BorderColor = ColorTranslator.FromHtml(borderColor.ToObject<string>());
            }

            var borderStyle = prop.SelectToken(nameof(Style.BorderStyle));
            if (borderStyle != null)
            {
                result.BorderStyle = borderStyle.ToObject<BorderStyle>();
            }

            var contentAlignment = prop.SelectToken(nameof(Style.ContentAlignment));
            if (contentAlignment != null)
            {
                result.ContentAlignment = contentAlignment.ToObject<ContentAlignment>();
            }

            var font = prop.SelectToken(nameof(Style.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    result.Font = fontObj;
                }
            }

            var foreColor = prop.SelectToken(nameof(Style.ForeColor));
            if (foreColor != null)
            {
                result.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var textEffect = prop.SelectToken(nameof(Style.TextEffect));
            if (textEffect != null)
            {
                result.TextEffect = textEffect.ToObject<TextEffect>();
            }

            return result;
        }
    }
}
