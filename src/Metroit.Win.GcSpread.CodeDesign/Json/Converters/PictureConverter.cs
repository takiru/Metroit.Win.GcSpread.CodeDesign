using FarPoint.Win;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// Picture オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class PictureConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="picture">Picture オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(Picture picture)
        {
            if (picture == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(Picture.AlignHorz), picture.AlignHorz));
            jObj.Add(new JProperty(nameof(Picture.AlignVert), picture.AlignVert));

            ImageConverter imgConverter = new ImageConverter();
            byte[] imageBytes = (byte[])imgConverter.ConvertTo(picture.Image, typeof(byte[]));
            jObj.Add(new JProperty(nameof(Picture.Image), System.Convert.ToBase64String(imageBytes)));

            jObj.Add(new JProperty(nameof(Picture.Style), picture.Style));
            jObj.Add(new JProperty(nameof(Picture.TransparencyColor), ColorTranslator.ToHtml(picture.TransparencyColor)));
            jObj.Add(new JProperty(nameof(Picture.TransparencyTolerance), picture.TransparencyTolerance));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>Picture オブジェクト。</returns>
        public static Picture Deserialize(JToken prop)
        {
            return Deserialize(prop, new Picture());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする Picture オブジェクト。</param>
        /// <returns>Picture オブジェクト。</returns>
        public static Picture Deserialize(JToken prop, Picture source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new Picture();
            if (source != null)
            {
                result.AlignHorz = source.AlignHorz;
                result.AlignVert = source.AlignVert;
                result.Image = source.Image;
                result.Style = source.Style;
                result.TransparencyColor = source.TransparencyColor;
                result.TransparencyTolerance = source.TransparencyTolerance;
            }

            var alignHorz = prop.SelectToken(nameof(Picture.AlignHorz));
            if (alignHorz != null)
            {
                result.AlignHorz = alignHorz.ToObject<HorizontalAlignment>();
            }

            var alignVert = prop.SelectToken(nameof(Picture.AlignVert));
            if (alignVert != null)
            {
                result.AlignVert = alignVert.ToObject<VerticalAlignment>();
            }

            var image = prop.SelectToken(nameof(Picture.Image));
            if (image != null)
            {
                var imageValue = System.Convert.FromBase64String(image.ToObject<string>());
                ImageConverter imgConverter = new ImageConverter();
                var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                result.Image = imageObj;
            }

            var style = prop.SelectToken(nameof(Picture.Style));
            if (style != null)
            {
                result.Style = style.ToObject<RenderStyle>();
            }

            var transparencyColor = prop.SelectToken(nameof(Picture.TransparencyColor));
            if (transparencyColor != null)
            {
                result.TransparencyColor = ColorTranslator.FromHtml(style.ToObject<string>());
            }

            var transparencyTolerance = prop.SelectToken(nameof(Picture.TransparencyTolerance));
            if (transparencyTolerance != null)
            {
                result.TransparencyTolerance = transparencyTolerance.ToObject<int>();
            }

            return result;
        }
    }
}
