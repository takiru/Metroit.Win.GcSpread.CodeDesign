using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;

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
            var allowSpin = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.BackColor));
            if (allowSpin != null)
            {
                dateSpin.AllowSpin = allowSpin.ToObject<bool>();
            }

            var increment = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.BorderStyle));
            if (increment != null)
            {
                dateSpin.Increment = increment.ToObject<int>();
            }

            var incrementValue = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.ButtonBackColor));
            if (incrementValue != null)
            {
                dateSpin.IncrementValue = incrementValue.ToObject<TimeSpan>();
            }

            var spinMode = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.ButtonForeColor));
            if (spinMode != null)
            {
                dateSpin.SpinMode = spinMode.ToObject<DateSpinMode>();
            }

            var spinOnKeys = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.SpinOnKeys));
            if (spinOnKeys != null)
            {
                dateSpin.SpinOnKeys = spinOnKeys.ToObject<bool>();
            }

            var spinOnWheel = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.SpinOnWheel));
            if (spinOnWheel != null)
            {
                dateSpin.SpinOnWheel = spinOnWheel.ToObject<bool>();
            }

            var wrap = prop.SelectToken(nameof(DateTimeDropDownPickerInfo.Wrap));
            if (wrap != null)
            {
                dateSpin.Wrap = wrap.ToObject<bool>();
            }
        }
    }
}
