using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// GradientEffect オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class GradientEffectConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="gradientEffect">GradientEffect オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(GradientEffect gradientEffect)
        {
            if (gradientEffect == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(GradientEffect.Direction), gradientEffect.Direction));
            jObj.Add(new JProperty(nameof(GradientEffect.EndColor), ColorTranslator.ToHtml(gradientEffect.EndColor)));
            jObj.Add(new JProperty(nameof(GradientEffect.StartColor), ColorTranslator.ToHtml(gradientEffect.StartColor)));
            jObj.Add(new JProperty(nameof(GradientEffect.Style), gradientEffect.Style));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>GradientEffect オブジェクト。</returns>
        public static GradientEffect Deserialize(JToken prop)
        {
            return Deserialize(prop, new GradientEffect());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする GradientEffect オブジェクト。</param>
        /// <returns>GradientEffect オブジェクト。</returns>
        public static GradientEffect Deserialize(JToken prop, GradientEffect source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new GradientEffect();
            if (source != null)
            {
                result.Direction = source.Direction;
                result.EndColor = source.EndColor;
                result.StartColor = source.StartColor;
                result.Style = source.Style;
            }

            var direction = prop.SelectToken(nameof(GradientEffect.Direction));
            if (direction != null)
            {
                result.Direction = direction.ToObject<GradientDirection>();
            }

            var endColor = prop.SelectToken(nameof(GradientEffect.EndColor));
            if (endColor != null)
            {
                result.EndColor = ColorTranslator.FromHtml(endColor.ToObject<string>());
            }

            var startColor = prop.SelectToken(nameof(GradientEffect.StartColor));
            if (startColor != null)
            {
                result.StartColor = ColorTranslator.FromHtml(startColor.ToObject<string>());
            }

            var style = prop.SelectToken(nameof(GradientEffect.Style));
            if (style != null)
            {
                result.Style = style.ToObject<GradientStyle>();
            }

            return result;
        }
    }
}
