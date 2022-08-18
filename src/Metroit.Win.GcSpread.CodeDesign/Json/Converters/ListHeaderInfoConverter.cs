using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ListHeaderInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ListHeaderInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="listHeaderInfo">ListHeaderInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(ListHeaderInfo listHeaderInfo)
        {
            if (listHeaderInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(ListHeaderInfo.AllowResize), listHeaderInfo.AllowResize));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.BackColor), ColorTranslator.ToHtml(listHeaderInfo.BackColor)));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.Clickable), listHeaderInfo.Clickable));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.ContentAlignment), listHeaderInfo.ContentAlignment));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.Ellipsis), listHeaderInfo.Ellipsis));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.ForeColor), ColorTranslator.ToHtml(listHeaderInfo.ForeColor)));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.GradientEffect), GradientEffectConverter.Serialize(listHeaderInfo.GradientEffect)));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.Image), listHeaderInfo.Image));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.ImageTextSpace), listHeaderInfo.ImageTextSpace));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.SortIndicatorAlignment), listHeaderInfo.SortIndicatorAlignment));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.Text), listHeaderInfo.Text));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.TextAttachAlignment), listHeaderInfo.TextAttachAlignment));
            jObj.Add(new JProperty(nameof(ListHeaderInfo.TextEffect), listHeaderInfo.TextEffect));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>ListHeaderInfo オブジェクト。</returns>
        public static ListHeaderInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new ListHeaderInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする ListHeaderInfo オブジェクト。</param>
        /// <returns>ListHeaderInfo オブジェクト。</returns>
        public static ListHeaderInfo Deserialize(JToken prop, ListHeaderInfo source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new ListHeaderInfo();
            if (source != null)
            {
                result.AllowResize = source.AllowResize;
                result.BackColor = source.BackColor;
                result.Clickable = source.Clickable;
                result.ContentAlignment = source.ContentAlignment;
                result.Ellipsis = source.Ellipsis;
                result.ForeColor = source.ForeColor;
                result.GradientEffect = source.GradientEffect;
                result.Image = source.Image;
                result.ImageTextSpace = source.ImageTextSpace;
                result.SortIndicatorAlignment = source.SortIndicatorAlignment;
                result.Text = source.Text;
                result.TextAttachAlignment = source.TextAttachAlignment;
                result.TextEffect = source.TextEffect;
            }

            var allowResize = prop.SelectToken(nameof(ListHeaderInfo.AllowResize));
            if (allowResize != null)
            {
                result.AllowResize = allowResize.ToObject<bool>();
            }

            var backColor = prop.SelectToken(nameof(ListHeaderInfo.BackColor));
            if (backColor != null)
            {
                result.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var clickable = prop.SelectToken(nameof(ListHeaderInfo.Clickable));
            if (clickable != null)
            {
                result.Clickable = clickable.ToObject<bool>();
            }

            var contentAlignment = prop.SelectToken(nameof(ListHeaderInfo.ContentAlignment));
            if (contentAlignment != null)
            {
                result.ContentAlignment = contentAlignment.ToObject<ContentAlignment>();
            }

            var ellipsis = prop.SelectToken(nameof(ListHeaderInfo.Ellipsis));
            if (ellipsis != null)
            {
                result.Ellipsis = contentAlignment.ToObject<bool>();
            }

            var foreColor = prop.SelectToken(nameof(ListHeaderInfo.ForeColor));
            if (foreColor != null)
            {
                result.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var gradientEffect = prop.SelectToken(nameof(ListHeaderInfo.GradientEffect));
            if (gradientEffect != null)
            {
                result.GradientEffect = GradientEffectConverter.Deserialize(gradientEffect, result.GradientEffect);
            }

            return result;
        }
    }
}
