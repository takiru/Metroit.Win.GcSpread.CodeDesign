using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DateTimeDropDownPickerInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DateTimeDropDownPickerInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dateTimeDropDownPickerInfo">DateTimeDropDownPickerInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DateTimeDropDownPickerInfo dateTimeDropDownPickerInfo)
        {
            if (dateTimeDropDownPickerInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.BackColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.BackColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.BorderStyle), dateTimeDropDownPickerInfo.BorderStyle));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.ButtonBackColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.ButtonBackColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.ButtonForeColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.ButtonForeColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.DateTabText), dateTimeDropDownPickerInfo.DateTabText));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.FlatStyle), dateTimeDropDownPickerInfo.FlatStyle));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.Font), FontConverter.Serialize(dateTimeDropDownPickerInfo.Font)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.ForeColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.ForeColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.HeaderBackColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.HeaderBackColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.SelectedBackColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.SelectedBackColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.SelectedBorderColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.SelectedBorderColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.SelectedForeColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.SelectedForeColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.SelectionRenderMode), dateTimeDropDownPickerInfo.SelectionRenderMode));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.ShowPickers), dateTimeDropDownPickerInfo.ShowPickers));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.SingleBorderColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.SingleBorderColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.TabForeColor), ColorTranslator.ToHtml(dateTimeDropDownPickerInfo.TabForeColor)));
            jObj.Add(new JProperty(nameof(DateTimeDropDownPickerInfo.TimeTabText), dateTimeDropDownPickerInfo.TimeTabText));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="dateTimeDropDownPickerInfo">DateTimeDropDownPickerInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(DateTimeDropDownPickerInfo dateTimeDropDownPickerInfo, JToken prop)
        {
            var backColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.BackColor));
            if (backColor != null)
            {
                dateTimeDropDownPickerInfo.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var borderStyle = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.BorderStyle));
            if (borderStyle != null)
            {
                dateTimeDropDownPickerInfo.BorderStyle = borderStyle.ToObject<BorderStyle>();
            }

            var buttonBackColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.ButtonBackColor));
            if (buttonBackColor != null)
            {
                dateTimeDropDownPickerInfo.ButtonBackColor = ColorTranslator.FromHtml(buttonBackColor.ToObject<string>());
            }

            var buttonForeColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.ButtonForeColor));
            if (buttonForeColor != null)
            {
                dateTimeDropDownPickerInfo.ButtonForeColor = ColorTranslator.FromHtml(buttonForeColor.ToObject<string>());
            }

            var dateTabText = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.DateTabText));
            if (dateTabText != null)
            {
                dateTimeDropDownPickerInfo.DateTabText = dateTabText.ToObject<string>();
            }

            var flatStyle = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.FlatStyle));
            if (flatStyle != null)
            {
                dateTimeDropDownPickerInfo.FlatStyle = flatStyle.ToObject<FlatStyle>();
            }

            var font = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    dateTimeDropDownPickerInfo.Font = fontObj;
                }
            }

            var foreColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.ForeColor));
            if (foreColor != null)
            {
                dateTimeDropDownPickerInfo.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var headerBackColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.HeaderBackColor));
            if (headerBackColor != null)
            {
                dateTimeDropDownPickerInfo.HeaderBackColor = ColorTranslator.FromHtml(headerBackColor.ToObject<string>());
            }

            var selectedBackColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.SelectedBackColor));
            if (selectedBackColor != null)
            {
                dateTimeDropDownPickerInfo.SelectedBackColor = ColorTranslator.FromHtml(selectedBackColor.ToObject<string>());
            }

            var selectedBorderColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.SelectedBorderColor));
            if (selectedBorderColor != null)
            {
                dateTimeDropDownPickerInfo.SelectedBorderColor = ColorTranslator.FromHtml(selectedBorderColor.ToObject<string>());
            }

            var selectedForeColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.SelectedForeColor));
            if (selectedForeColor != null)
            {
                dateTimeDropDownPickerInfo.SelectedForeColor = ColorTranslator.FromHtml(selectedForeColor.ToObject<string>());
            }

            var selectionRenderMode = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.SelectionRenderMode));
            if (selectionRenderMode != null)
            {
                dateTimeDropDownPickerInfo.SelectionRenderMode = selectionRenderMode.ToObject<SelectionRenderMode>();
            }

            var showPickers = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.ShowPickers));
            if (showPickers != null)
            {
                dateTimeDropDownPickerInfo.ShowPickers = showPickers.ToObject<PickerDisplayOptions>();
            }

            var singleBorderColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.SingleBorderColor));
            if (singleBorderColor != null)
            {
                dateTimeDropDownPickerInfo.SingleBorderColor = ColorTranslator.FromHtml(singleBorderColor.ToObject<string>());
            }

            var tabForeColor = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.TabForeColor));
            if (tabForeColor != null)
            {
                dateTimeDropDownPickerInfo.TabForeColor = ColorTranslator.FromHtml(tabForeColor.ToObject<string>());
            }

            var timeTabText = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.TimeTabText));
            if (timeTabText != null)
            {
                dateTimeDropDownPickerInfo.TimeTabText = timeTabText.ToObject<string>();
            }
        }
    }
}
