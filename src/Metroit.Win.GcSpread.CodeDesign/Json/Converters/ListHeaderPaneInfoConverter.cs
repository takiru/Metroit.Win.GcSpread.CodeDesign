using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ListHeaderPaneInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ListHeaderPaneInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="listHeaderPaneInfo">ListHeaderPaneInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(ListHeaderPaneInfo listHeaderPaneInfo)
        {
            if (listHeaderPaneInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(ListHeaderPaneInfo.AutoHeight), listHeaderPaneInfo.AutoHeight));
            jObj.Add(new JProperty(nameof(ListHeaderPaneInfo.BackColor), ColorTranslator.ToHtml(listHeaderPaneInfo.BackColor)));
            jObj.Add(new JProperty(nameof(ListHeaderPaneInfo.Font), FontConverter.Serialize(listHeaderPaneInfo.Font)));
            jObj.Add(new JProperty(nameof(ListHeaderPaneInfo.GradientEffect), GradientEffectConverter.Serialize(listHeaderPaneInfo.GradientEffect)));
            jObj.Add(new JProperty(nameof(ListHeaderPaneInfo.Height), listHeaderPaneInfo.Height));
            jObj.Add(new JProperty(nameof(ListHeaderPaneInfo.Visible), listHeaderPaneInfo.Visible));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="listHeaderPaneInfo">ListHeaderPaneInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(ListHeaderPaneInfo listHeaderPaneInfo, JToken prop)
        {
            var autoHeight = prop.SelectToken(nameof(ListHeaderPaneInfo.AutoHeight));
            if (autoHeight != null)
            {
                listHeaderPaneInfo.AutoHeight = autoHeight.ToObject<bool>();
            }

            var backColor = prop.SelectToken(nameof(ListHeaderPaneInfo.BackColor));
            if (backColor != null)
            {
                listHeaderPaneInfo.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var font = prop.SelectToken(nameof(ListHeaderPaneInfo.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    listHeaderPaneInfo.Font = fontObj;
                }
            }

            var gradientEffect = prop.SelectToken(nameof(ListHeaderPaneInfo.GradientEffect));
            if (gradientEffect != null)
            {
                listHeaderPaneInfo.GradientEffect = GradientEffectConverter.Deserialize(gradientEffect, listHeaderPaneInfo.GradientEffect);
            }

            var height = prop.SelectToken(nameof(ListHeaderPaneInfo.Height));
            if (height != null)
            {
                listHeaderPaneInfo.Height = height.ToObject<int>();
            }

            var visible = prop.SelectToken(nameof(ListHeaderPaneInfo.Visible));
            if (visible != null)
            {
                listHeaderPaneInfo.Visible = visible.ToObject<bool>();
            }
        }
    }
}
