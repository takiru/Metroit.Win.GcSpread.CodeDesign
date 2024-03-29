﻿using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ComboDropDownInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ComboDropDownInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="comboDropDownInfo">ComboDropDownInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(ComboDropDownInfo comboDropDownInfo)
        {
            if (comboDropDownInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(ComboDropDownInfo.AllowDrop), comboDropDownInfo.AllowDrop));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.AllowResize), comboDropDownInfo.AllowResize));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.AutoDropDown), comboDropDownInfo.AutoDropDown));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.AutoHideTouchKeyboard), comboDropDownInfo.AutoHideTouchKeyboard));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.AutoWidth), comboDropDownInfo.AutoWidth));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.ClosingAnimation), comboDropDownInfo.ClosingAnimation));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.Direction), comboDropDownInfo.Direction));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.Height), comboDropDownInfo.Height));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.OpeningAnimation), comboDropDownInfo.OpeningAnimation));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.ShowShadow), comboDropDownInfo.ShowShadow));
            jObj.Add(new JProperty(nameof(ComboDropDownInfo.Width), comboDropDownInfo.Width));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="comboDropDownInfo">ComboDropDownInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(ComboDropDownInfo comboDropDownInfo, JToken prop)
        {
            var allowDrop = prop.SelectToken(nameof(ComboDropDownInfo.AllowDrop));
            if (allowDrop != null)
            {
                comboDropDownInfo.AllowDrop = allowDrop.ToObject<bool>();
            }

            var allowResize = prop.SelectToken(nameof(ComboDropDownInfo.AllowResize));
            if (allowResize != null)
            {
                comboDropDownInfo.AllowResize = allowResize.ToObject<bool>();
            }

            var autoDropDown = prop.SelectToken(nameof(ComboDropDownInfo.AutoDropDown));
            if (autoDropDown != null)
            {
                comboDropDownInfo.AutoDropDown = autoDropDown.ToObject<bool>();
            }

            var autoHideTouchKeyboard = prop.SelectToken(nameof(ComboDropDownInfo.AutoHideTouchKeyboard));
            if (autoHideTouchKeyboard != null)
            {
                comboDropDownInfo.AutoHideTouchKeyboard = autoHideTouchKeyboard.ToObject<AutoHideTouchKeyboard>();
            }

            var closingAnimation = prop.SelectToken(nameof(ComboDropDownInfo.ClosingAnimation));
            if (closingAnimation != null)
            {
                comboDropDownInfo.ClosingAnimation = closingAnimation.ToObject<DropDownAnimation>();
            }

            var direction = prop.SelectToken(nameof(ComboDropDownInfo.Direction));
            if (direction != null)
            {
                comboDropDownInfo.Direction = direction.ToObject<DropDownDirection>();
            }

            var height = prop.SelectToken(nameof(ComboDropDownInfo.Height));
            if (height != null)
            {
                comboDropDownInfo.Height = height.ToObject<int>();
            }

            var openingAnimation = prop.SelectToken(nameof(ComboDropDownInfo.OpeningAnimation));
            if (openingAnimation != null)
            {
                comboDropDownInfo.OpeningAnimation = openingAnimation.ToObject<DropDownAnimation>();
            }

            var showShadow = prop.SelectToken(nameof(ComboDropDownInfo.ShowShadow));
            if (showShadow != null)
            {
                comboDropDownInfo.ShowShadow = showShadow.ToObject<bool>();
            }

            var width = prop.SelectToken(nameof(ComboDropDownInfo.Width));
            if (width != null)
            {
                comboDropDownInfo.Width = width.ToObject<int>();
            }

            // NOTE: Width より後に設定しないと自動調整が効かない
            var autoWidth = prop.SelectToken(nameof(ComboDropDownInfo.AutoWidth));
            if (autoWidth != null)
            {
                comboDropDownInfo.AutoWidth = autoWidth.ToObject<bool>();
            }
        }
    }
}
