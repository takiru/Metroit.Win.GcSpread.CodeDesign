using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DateTimeDropDownInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DateTimeDropDownInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dateTimeDropDownInfo">DateTimeDropDownInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DateTimeDropDownInfo dateTimeDropDownInfo)
        {
            if (dateTimeDropDownInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.AllowDrop), dateTimeDropDownInfo.AllowDrop));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.AllowResize), dateTimeDropDownInfo.AllowResize));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.AutoDropDown), dateTimeDropDownInfo.AutoDropDown));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.AutoHideTouchKeyboard), dateTimeDropDownInfo.AutoHideTouchKeyboard));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.ClosingAnimation), dateTimeDropDownInfo.ClosingAnimation));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.Direction), dateTimeDropDownInfo.Direction));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.DropDownType), dateTimeDropDownInfo.DropDownType));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.OpeningAnimation), dateTimeDropDownInfo.OpeningAnimation));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.ShowShadow), dateTimeDropDownInfo.ShowShadow));
            jObj.Add(new JProperty(nameof(DateTimeDropDownInfo.Size), $"{dateTimeDropDownInfo.Size.Width}, {dateTimeDropDownInfo.Size.Height}"));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="dateTimeDropDownInfo">DateTimeDropDownInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(DateTimeDropDownInfo dateTimeDropDownInfo, JToken prop)
        {
            var allowDrop = prop.SelectToken(nameof(DateTimeDropDownInfo.AllowDrop));
            if (allowDrop != null)
            {
                dateTimeDropDownInfo.AllowDrop = allowDrop.ToObject<bool>();
            }

            var allowResize = prop.SelectToken(nameof(DateTimeDropDownInfo.AllowResize));
            if (allowResize != null)
            {
                dateTimeDropDownInfo.AllowResize = allowResize.ToObject<bool>();
            }

            var autoDropDown = prop.SelectToken(nameof(DateTimeDropDownInfo.AutoDropDown));
            if (autoDropDown != null)
            {
                dateTimeDropDownInfo.AutoDropDown = autoDropDown.ToObject<bool>();
            }

            var autoHideTouchKeyboard = prop.SelectToken(nameof(DateTimeDropDownInfo.AutoHideTouchKeyboard));
            if (autoHideTouchKeyboard != null)
            {
                dateTimeDropDownInfo.AutoHideTouchKeyboard = autoHideTouchKeyboard.ToObject<AutoHideTouchKeyboard>();
            }

            var closingAnimation = prop.SelectToken(nameof(DateTimeDropDownInfo.ClosingAnimation));
            if (closingAnimation != null)
            {
                dateTimeDropDownInfo.ClosingAnimation = closingAnimation.ToObject<DropDownAnimation>();
            }

            var direction = prop.SelectToken(nameof(DateTimeDropDownInfo.Direction));
            if (direction != null)
            {
                dateTimeDropDownInfo.Direction = direction.ToObject<DropDownDirection>();
            }

            var dropDownType = prop.SelectToken(nameof(DateTimeDropDownInfo.DropDownType));
            if (dropDownType != null)
            {
                dateTimeDropDownInfo.DropDownType = dropDownType.ToObject<DateDropDownType>();
            }

            var openingAnimation = prop.SelectToken(nameof(DateTimeDropDownInfo.OpeningAnimation));
            if (openingAnimation != null)
            {
                dateTimeDropDownInfo.OpeningAnimation = openingAnimation.ToObject<DropDownAnimation>();
            }

            var showShadow = prop.SelectToken(nameof(DateTimeDropDownInfo.ShowShadow));
            if (showShadow != null)
            {
                dateTimeDropDownInfo.ShowShadow = showShadow.ToObject<bool>();
            }

            var size = prop.SelectToken(nameof(DateTimeDropDownInfo.Size));
            if (size != null)
            {
                dateTimeDropDownInfo.Size = size.ToObject<Size>();
            }
        }
    }
}
