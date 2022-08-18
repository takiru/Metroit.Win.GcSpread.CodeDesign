using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DropDownEditorInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DropDownEditorInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// TouchToolBar はシリアライズを行いません。
        /// </summary>
        /// <param name="dropDownInfo">DropDownInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         TouchToolBar: 大量のプロパティがあるため。
        public static JObject Serialize(DropDownEditorInfo dropDownEditorInfo)
        {
            if (dropDownEditorInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DropDownEditorInfo.BackColor), ColorTranslator.ToHtml(dropDownEditorInfo.BackColor)));

            ImageConverter imgConverter = new ImageConverter();
            byte[] imageBytes = (byte[])imgConverter.ConvertTo(dropDownEditorInfo.BackgroundImage, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(DropDownEditorInfo.BackgroundImage), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(DropDownEditorInfo.BackgroundImage), System.Convert.ToBase64String(imageBytes)));
            }

            jObj.Add(new JProperty(nameof(DropDownEditorInfo.BackgroundImageLayout), dropDownEditorInfo.BackgroundImageLayout));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.BorderStyle), dropDownEditorInfo.BorderStyle));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.ContentAlignment), dropDownEditorInfo.ContentAlignment));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.Font), FontConverter.Serialize(dropDownEditorInfo.Font)));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.ForeColor), ColorTranslator.ToHtml(dropDownEditorInfo.ForeColor)));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.Padding),
                $"{dropDownEditorInfo.Padding.Left}, {dropDownEditorInfo.Padding.Top}, {dropDownEditorInfo.Padding.Right}, {dropDownEditorInfo.Padding.Bottom}"));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.ReadOnly), dropDownEditorInfo.ReadOnly));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.ScrollBarMode), dropDownEditorInfo.ScrollBarMode));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.ScrollBars), dropDownEditorInfo.ScrollBars));
            jObj.Add(new JProperty(nameof(DropDownEditorInfo.SingleBorderColor), ColorTranslator.ToHtml(dropDownEditorInfo.SingleBorderColor)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// TouchToolBar はデシリアライズを行いません。
        /// </summary>
        /// <param name="dropDownInfo">DropDownInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        // NOTE: 下記はシリアライズ化対象外。
        //         TouchToolBar: 大量のプロパティがあるため。
        public static void Deserialize(DropDownEditorInfo dropDownEditorInfo, JToken prop)
        {
            var backColor = prop.SelectToken(nameof(DropDownEditorInfo.BackColor));
            if (backColor != null)
            {
                dropDownEditorInfo.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var backgroundImage = prop.SelectToken(nameof(DropDownEditorInfo.BackgroundImage));
            if (backgroundImage != null)
            {
                if (backgroundImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(backgroundImage.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    dropDownEditorInfo.BackgroundImage = imageObj;
                }
                else
                {
                    dropDownEditorInfo.BackgroundImage = null;
                }
            }

            var backgroundImageLayout = prop.SelectToken(nameof(DropDownEditorInfo.BackgroundImageLayout));
            if (backgroundImageLayout != null)
            {
                dropDownEditorInfo.BackgroundImageLayout = backgroundImageLayout.ToObject<ImageLayout>();
            }

            var borderStyle = prop.SelectToken(nameof(DropDownEditorInfo.BorderStyle));
            if (borderStyle != null)
            {
                dropDownEditorInfo.BorderStyle = borderStyle.ToObject<BorderStyle>();
            }

            var contentAlignment = prop.SelectToken(nameof(DropDownEditorInfo.ContentAlignment));
            if (contentAlignment != null)
            {
                dropDownEditorInfo.ContentAlignment = contentAlignment.ToObject<ContentAlignment>();
            }

            var font = prop.SelectToken(nameof(DropDownEditorInfo.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    dropDownEditorInfo.Font = fontObj;
                }
            }

            var foreColor = prop.SelectToken(nameof(DropDownEditorInfo.ForeColor));
            if (foreColor != null)
            {
                dropDownEditorInfo.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var padding = prop.SelectToken(nameof(DropDownEditorInfo.Padding));
            if (padding != null)
            {
                dropDownEditorInfo.Padding = padding.ToObject<Padding>();
            }

            var readOnly = prop.SelectToken(nameof(DropDownEditorInfo.ReadOnly));
            if (readOnly != null)
            {
                dropDownEditorInfo.ReadOnly = readOnly.ToObject<bool>();
            }

            var scrollBarMode = prop.SelectToken(nameof(DropDownEditorInfo.ScrollBarMode));
            if (scrollBarMode != null)
            {
                dropDownEditorInfo.ScrollBarMode = scrollBarMode.ToObject<ScrollBarMode>();
            }

            var scrollBars = prop.SelectToken(nameof(DropDownEditorInfo.ScrollBars));
            if (scrollBars != null)
            {
                dropDownEditorInfo.ScrollBars = scrollBars.ToObject<ScrollBars>();
            }

            var singleBorderColor = prop.SelectToken(nameof(DropDownEditorInfo.SingleBorderColor));
            if (singleBorderColor != null)
            {
                dropDownEditorInfo.SingleBorderColor = ColorTranslator.FromHtml(singleBorderColor.ToObject<string>());
            }
        }
    }
}
