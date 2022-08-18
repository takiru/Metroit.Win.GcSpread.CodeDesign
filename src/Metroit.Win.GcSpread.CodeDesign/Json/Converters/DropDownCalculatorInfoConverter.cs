using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DropDownCalculatorInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DropDownCalculatorInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// ButtonKeys, ButtonText はシリアライズを行いません。
        /// </summary>
        /// <param name="dropDownCalculatorInfo">DropDownCalculatorInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         ButtonKeys: 固定情報である、デザイナで指定不可なため。
        //         ButtonText: デザイナで指定不可なため。
        public static JObject Serialize(DropDownCalculatorInfo dropDownCalculatorInfo)
        {
            if (dropDownCalculatorInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.BackColor), ColorTranslator.ToHtml(dropDownCalculatorInfo.BackColor)));

            ImageConverter imgConverter = new ImageConverter();
            byte[] imageBytes = (byte[])imgConverter.ConvertTo(dropDownCalculatorInfo.BackgroundImage, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.BackgroundImage), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.BackgroundImage), System.Convert.ToBase64String(imageBytes)));
            }

            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.BackgroundImageLayout), dropDownCalculatorInfo.BackgroundImageLayout));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.BorderStyle), dropDownCalculatorInfo.BorderStyle));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.ContentAlignment), dropDownCalculatorInfo.ContentAlignment));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.EditButtons), CalculatorButtonStyleInfoConverter.Serialize(dropDownCalculatorInfo.EditButtons)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.FlatStyle), dropDownCalculatorInfo.FlatStyle));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.Font), FontConverter.Serialize(dropDownCalculatorInfo.Font)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.MathButtons), CalculatorButtonStyleInfoConverter.Serialize(dropDownCalculatorInfo.MathButtons)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.MemoryButtons), CalculatorButtonStyleInfoConverter.Serialize(dropDownCalculatorInfo.MemoryButtons)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.MemoryStatus), CalculatorButtonStyleInfoConverter.Serialize(dropDownCalculatorInfo.MemoryStatus)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.NumericButtons), CalculatorButtonStyleInfoConverter.Serialize(dropDownCalculatorInfo.NumericButtons)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.Output), CalculatorButtonStyleInfoConverter.Serialize(dropDownCalculatorInfo.Output)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.OutputHeight), dropDownCalculatorInfo.OutputHeight));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.Padding),
                $"{dropDownCalculatorInfo.Padding.Left}, {dropDownCalculatorInfo.Padding.Top}, {dropDownCalculatorInfo.Padding.Right}, {dropDownCalculatorInfo.Padding.Bottom}"));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.ResetButtons), CalculatorButtonStyleInfoConverter.Serialize(dropDownCalculatorInfo.ResetButtons)));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.ShowOutput), dropDownCalculatorInfo.ShowOutput));
            jObj.Add(new JProperty(nameof(DropDownCalculatorInfo.SingleBorderColor), ColorTranslator.ToHtml(dropDownCalculatorInfo.SingleBorderColor)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// ButtonKeys, ButtonText はデシリアライズを行いません。
        /// </summary>
        /// <param name="dropDownCalculatorInfo">DropDownCalculatorInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         ButtonKeys: 固定情報である、デザイナで指定不可なため。
        //         ButtonText: デザイナで指定不可なため。
        public static void Deserialize(DropDownCalculatorInfo dropDownCalculatorInfo, JToken prop)
        {
            var backColor = prop.SelectToken(nameof(DropDownCalculatorInfo.BackColor));
            if (backColor != null)
            {
                dropDownCalculatorInfo.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var backgroundImage = prop.SelectToken(nameof(DropDownCalculatorInfo.BackgroundImage));
            if (backgroundImage != null)
            {
                if (backgroundImage.HasValues)
                {
                    var imageValue = System.Convert.FromBase64String(backgroundImage.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    dropDownCalculatorInfo.BackgroundImage = imageObj;
                }
                else
                {
                    dropDownCalculatorInfo.BackgroundImage = null;
                }
            }

            var backgroundImageLayout = prop.SelectToken(nameof(DropDownCalculatorInfo.BackgroundImageLayout));
            if (backgroundImageLayout != null)
            {
                dropDownCalculatorInfo.BackgroundImageLayout = backgroundImageLayout.ToObject<ImageLayout>();
            }

            var borderStyle = prop.SelectToken(nameof(DropDownCalculatorInfo.BorderStyle));
            if (borderStyle != null)
            {
                dropDownCalculatorInfo.BorderStyle = borderStyle.ToObject<BorderStyle>();
            }

            var contentAlignment = prop.SelectToken(nameof(DropDownCalculatorInfo.ContentAlignment));
            if (contentAlignment != null)
            {
                dropDownCalculatorInfo.ContentAlignment = contentAlignment.ToObject<ContentAlignment>();
            }

            var editButtons = prop.SelectToken(nameof(DropDownCalculatorInfo.EditButtons));
            if (editButtons != null)
            {
                dropDownCalculatorInfo.EditButtons = CalculatorButtonStyleInfoConverter.Deserialize(editButtons, dropDownCalculatorInfo.EditButtons);
            }

            var flatStyle = prop.SelectToken(nameof(DropDownCalculatorInfo.FlatStyle));
            if (flatStyle != null)
            {
                dropDownCalculatorInfo.FlatStyle = flatStyle.ToObject<FlatStyle>();
            }

            var font = prop.SelectToken(nameof(DropDownCalculatorInfo.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    dropDownCalculatorInfo.Font = fontObj;
                }
            }

            var mathButtons = prop.SelectToken(nameof(DropDownCalculatorInfo.MathButtons));
            if (mathButtons != null)
            {
                dropDownCalculatorInfo.MathButtons = CalculatorButtonStyleInfoConverter.Deserialize(mathButtons, dropDownCalculatorInfo.MathButtons);
            }

            var memoryButtons = prop.SelectToken(nameof(DropDownCalculatorInfo.MemoryButtons));
            if (memoryButtons != null)
            {
                dropDownCalculatorInfo.MemoryButtons = CalculatorButtonStyleInfoConverter.Deserialize(memoryButtons, dropDownCalculatorInfo.MemoryButtons);
            }

            var memoryStatus = prop.SelectToken(nameof(DropDownCalculatorInfo.MemoryStatus));
            if (memoryStatus != null)
            {
                dropDownCalculatorInfo.MemoryStatus = CalculatorButtonStyleInfoConverter.Deserialize(memoryStatus, dropDownCalculatorInfo.MemoryStatus);
            }

            var numericButtons = prop.SelectToken(nameof(DropDownCalculatorInfo.NumericButtons));
            if (numericButtons != null)
            {
                dropDownCalculatorInfo.NumericButtons = CalculatorButtonStyleInfoConverter.Deserialize(numericButtons, dropDownCalculatorInfo.NumericButtons);
            }

            var output = prop.SelectToken(nameof(DropDownCalculatorInfo.Output));
            if (output != null)
            {
                dropDownCalculatorInfo.Output = CalculatorButtonStyleInfoConverter.Deserialize(output, dropDownCalculatorInfo.Output);
            }

            var outputHeight = prop.SelectToken(nameof(DropDownCalculatorInfo.OutputHeight));
            if (outputHeight != null)
            {
                dropDownCalculatorInfo.OutputHeight = outputHeight.ToObject<int>();
            }

            var padding = prop.SelectToken(nameof(DropDownCalculatorInfo.Padding));
            if (padding != null)
            {
                dropDownCalculatorInfo.Padding = padding.ToObject<Padding>();
            }

            var resetButtons = prop.SelectToken(nameof(DropDownCalculatorInfo.ResetButtons));
            if (resetButtons != null)
            {
                dropDownCalculatorInfo.ResetButtons = CalculatorButtonStyleInfoConverter.Deserialize(resetButtons, dropDownCalculatorInfo.ResetButtons);
            }

            var showOutput = prop.SelectToken(nameof(DropDownCalculatorInfo.ShowOutput));
            if (showOutput != null)
            {
                dropDownCalculatorInfo.ShowOutput = showOutput.ToObject<bool>();
            }

            var singleBorderColor = prop.SelectToken(nameof(DropDownCalculatorInfo.SingleBorderColor));
            if (singleBorderColor != null)
            {
                dropDownCalculatorInfo.SingleBorderColor = ColorTranslator.FromHtml(singleBorderColor.ToObject<string>());
            }
        }
    }
}
