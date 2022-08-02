using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// PopUpCalculatorInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class PopUpCalculatorInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="popUpCalculatorInfo">PopUpCalculatorInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(PopUpCalculatorInfo popUpCalculatorInfo)
        {
            if (popUpCalculatorInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.Align), popUpCalculatorInfo.Align));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.AllowPopUp), popUpCalculatorInfo.AllowPopUp));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.BackColor), ColorTranslator.ToHtml(popUpCalculatorInfo.BackColor)));

            ImageConverter imgConverter = new ImageConverter();
            byte[] imageBytes = (byte[])imgConverter.ConvertTo(popUpCalculatorInfo.BackgroundImage, typeof(byte[]));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.BackgroundImage), System.Convert.ToBase64String(imageBytes)));

            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.BackgroundImageLayout), popUpCalculatorInfo.BackgroundImageLayout));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.BorderStyle), popUpCalculatorInfo.BorderStyle));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.DecimalDigit), popUpCalculatorInfo.DecimalDigit));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.Font), FontConverter.Serialize(popUpCalculatorInfo.Font)));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.ForeColor), ColorTranslator.ToHtml(popUpCalculatorInfo.ForeColor)));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.Lines), popUpCalculatorInfo.Lines));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.Padding), 
                $"{popUpCalculatorInfo.Padding.Left}, {popUpCalculatorInfo.Padding.Top}, {popUpCalculatorInfo.Padding.Right}, {popUpCalculatorInfo.Padding.Bottom}"));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.TextHAlign), popUpCalculatorInfo.TextHAlign));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.UseDecimalDigit), popUpCalculatorInfo.UseDecimalDigit));
            jObj.Add(new JProperty(nameof(PopUpCalculatorInfo.Width), popUpCalculatorInfo.Width));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="popUpCalculatorInfo">PopUpCalculatorInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(PopUpCalculatorInfo popUpCalculatorInfo, JToken prop)
        {
            var align = prop.SelectToken(nameof(PopUpCalculatorInfo.Align));
            if (align != null)
            {
                popUpCalculatorInfo.Align = align.ToObject<LeftRightAlignment>();
            }

            var allowPopUp = prop.SelectToken(nameof(PopUpCalculatorInfo.AllowPopUp));
            if (allowPopUp != null)
            {
                popUpCalculatorInfo.AllowPopUp = allowPopUp.ToObject<bool>();
            }

            var backColor = prop.SelectToken(nameof(PopUpCalculatorInfo.BackColor));
            if (backColor != null)
            {
                popUpCalculatorInfo.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var backgroundImage = prop.SelectToken(nameof(PopUpCalculatorInfo.BackgroundImage));
            if (backgroundImage != null)
            {
                if (!string.IsNullOrWhiteSpace( backgroundImage.ToObject<string>()))
                {
                    var imageValue = System.Convert.FromBase64String(backgroundImage.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = imgConverter.ConvertFrom(imageValue);
                    popUpCalculatorInfo.BackgroundImage = (Image)imageObj;
                }
            }

            var backgroundImageLayout = prop.SelectToken(nameof(PopUpCalculatorInfo.BackgroundImageLayout));
            if (backgroundImageLayout != null)
            {
                popUpCalculatorInfo.BackgroundImageLayout = backgroundImageLayout.ToObject<ImageLayout>();
            }

            var borderStyle = prop.SelectToken(nameof(PopUpCalculatorInfo.BorderStyle));
            if (borderStyle != null)
            {
                popUpCalculatorInfo.BorderStyle = borderStyle.ToObject<BorderStyle>();
            }

            var decimalDigit = prop.SelectToken(nameof(PopUpCalculatorInfo.DecimalDigit));
            if (decimalDigit != null)
            {
                popUpCalculatorInfo.DecimalDigit = decimalDigit.ToObject<int>();
            }

            var font = prop.SelectToken(nameof(PopUpCalculatorInfo.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    popUpCalculatorInfo.Font = fontObj;
                }
            }

            var foreColor = prop.SelectToken(nameof(PopUpCalculatorInfo.ForeColor));
            if (foreColor != null)
            {
                popUpCalculatorInfo.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var lines = prop.SelectToken(nameof(PopUpCalculatorInfo.Lines));
            if (lines != null)
            {
                popUpCalculatorInfo.Lines = lines.ToObject<int>();
            }

            var padding = prop.SelectToken(nameof(PopUpCalculatorInfo.Padding));
            if (padding != null)
            {
                popUpCalculatorInfo.Padding = padding.ToObject<Padding>();
            }

            var textHAlign = prop.SelectToken(nameof(PopUpCalculatorInfo.TextHAlign));
            if (textHAlign != null)
            {
                popUpCalculatorInfo.TextHAlign = textHAlign.ToObject<HorizontalAlignment>();
            }

            var useDecimalDigit = prop.SelectToken(nameof(PopUpCalculatorInfo.UseDecimalDigit));
            if (useDecimalDigit != null)
            {
                popUpCalculatorInfo.UseDecimalDigit = useDecimalDigit.ToObject<bool>();
            }

            var width = prop.SelectToken(nameof(PopUpCalculatorInfo.Width));
            if (width != null)
            {
                popUpCalculatorInfo.Width = width.ToObject<int>();
            }
        }
    }
}
