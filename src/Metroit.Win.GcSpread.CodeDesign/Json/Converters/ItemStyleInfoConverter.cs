using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ItemStyleInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ItemStyleInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="itemStyleInfo">ItemStyleInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(ItemStyleInfo itemStyleInfo)
        {
            if (itemStyleInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(itemStyleInfo.BackColor), ColorTranslator.ToHtml(itemStyleInfo.BackColor)));
            jObj.Add(new JProperty(nameof(itemStyleInfo.ForeColor), ColorTranslator.ToHtml(itemStyleInfo.ForeColor)));
            jObj.Add(new JProperty(nameof(itemStyleInfo.GradientEffect), GradientEffectConverter.Serialize(itemStyleInfo.GradientEffect)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="itemStyleInfo">ItemStyleInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(ItemStyleInfo itemStyleInfo, JToken prop)
        {
            var backColor = prop.SelectToken(nameof(ItemStyleInfo.BackColor));
            if (backColor != null)
            {
                itemStyleInfo.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var foreColor = prop.SelectToken(nameof(ItemStyleInfo.ForeColor));
            if (foreColor != null)
            {
                itemStyleInfo.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var gradientEffect = prop.SelectToken(nameof(ItemStyleInfo.GradientEffect));
            if (gradientEffect != null)
            {
                itemStyleInfo.GradientEffect = GradientEffectConverter.Deserialize(gradientEffect, itemStyleInfo.GradientEffect);
            }
        }
    }
}
