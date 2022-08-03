using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ItemTemplateCollectionInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ItemTemplateCollectionInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="itemTemplateCollectionInfo">ItemTemplateCollectionInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(ItemTemplateCollectionInfo itemTemplateCollectionInfo)
        {
            if (itemTemplateCollectionInfo == null)
            {
                return null;
            }

            var jArray = new JArray();

            foreach (ItemTemplateInfo iti in itemTemplateCollectionInfo)
            {
                if (iti == null)
                {
                    continue;
                }

                var jObj = new JObject();
                jObj.Add(new JProperty(nameof(ItemTemplateInfo.AutoItemHeight), iti.AutoItemHeight));
                jObj.Add(new JProperty(nameof(ItemTemplateInfo.BackColor), ColorTranslator.ToHtml(iti.BackColor)));
                jObj.Add(new JProperty(nameof(ItemTemplateInfo.Font), FontConverter.Serialize(iti.Font)));
                jObj.Add(new JProperty(nameof(ItemTemplateInfo.ForeColor), ColorTranslator.ToHtml(iti.ForeColor)));
                jObj.Add(new JProperty(nameof(ItemTemplateInfo.GradientEffect), GradientEffectConverter.Serialize(iti.GradientEffect)));
                jObj.Add(new JProperty(nameof(ItemTemplateInfo.Height), iti.Height));

                ImageConverter imgConverter = new ImageConverter();
                byte[] imageBytes = (byte[])imgConverter.ConvertTo(iti.Image, typeof(byte[]));
                jObj.Add(new JProperty(nameof(ItemTemplateInfo.Image), System.Convert.ToBase64String(imageBytes)));

                jObj.Add(new JProperty(nameof(ItemTemplateInfo.Indent), iti.Indent));

                jArray.Add(jObj);
            }

            return jArray;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <returns>ItemTemplateCollectionInfo オブジェクト。</returns>
        public static ItemTemplateCollectionInfo Deserialize(JArray props)
        {
            return Deserialize(props, new ItemTemplateCollectionInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="props">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする ItemTemplateCollectionInfo オブジェクト。</param>
        /// <returns>ItemTemplateCollectionInfo オブジェクト。</returns>
        public static ItemTemplateCollectionInfo Deserialize(JArray props, ItemTemplateCollectionInfo source)
        {
            if (props.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new ItemTemplateCollectionInfo();
            foreach (ItemTemplateInfo iti in source)
            {
                var newIti = new ItemTemplateInfo();
                newIti.AutoItemHeight = iti.AutoItemHeight;
                newIti.BackColor = iti.BackColor;
                newIti.Font = iti.Font;
                newIti.ForeColor = iti.ForeColor;
                newIti.GradientEffect = iti.GradientEffect;
                newIti.Height = iti.Height;
                newIti.Image = iti.Image;
                newIti.Indent = iti.Indent;
            }

            var i = 0;
            foreach (JToken prop in props)
            {
                if (result.Count - 1 < i)
                {
                    result.Add(new ItemTemplateInfo());
                }

                var autoItemHeight = prop.SelectToken(nameof(ItemTemplateInfo.AutoItemHeight));
                if (autoItemHeight != null)
                {
                    result[i].AutoItemHeight = autoItemHeight.ToObject<bool>();
                }

                var backColor = prop.SelectToken(nameof(ItemTemplateInfo.BackColor));
                if (backColor != null)
                {
                    result[i].BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
                }

                var font = prop.SelectToken(nameof(ItemTemplateInfo.Font));
                if (font != null)
                {
                    var fontObj = FontConverter.Deserialize(font);
                    if (fontObj != null)
                    {
                        result[i].Font = fontObj;
                    }
                }

                var foreColor = prop.SelectToken(nameof(ItemTemplateInfo.ForeColor));
                if (foreColor != null)
                {
                    result[i].ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
                }

                var gradientEffect = prop.SelectToken(nameof(ItemTemplateInfo.GradientEffect));
                if (gradientEffect != null)
                {
                    result[i].GradientEffect = GradientEffectConverter.Deserialize(gradientEffect, result[i].GradientEffect);
                }

                var height = prop.SelectToken(nameof(ItemTemplateInfo.Height));
                if (height != null)
                {
                    result[i].Height = height.ToObject<int>();
                }

                var image = prop.SelectToken(nameof(ItemTemplateInfo.Image));
                if (image != null)
                {
                    var imageValue = System.Convert.FromBase64String(image.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = imgConverter.ConvertFrom(imageValue);
                    result[i].Image = imageObj;
                }

                var indent = prop.SelectToken(nameof(ItemTemplateInfo.Indent));
                if (indent != null)
                {
                    result[i].Indent = indent.ToObject<int>();
                }

                i++;
            }

            return result;
        }
    }
}
