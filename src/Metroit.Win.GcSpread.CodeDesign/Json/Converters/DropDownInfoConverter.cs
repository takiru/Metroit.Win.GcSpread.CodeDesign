using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DropDownInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DropDownInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dropDownInfo">DropDownInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DropDownInfo dropDownInfo)
        {
            if (dropDownInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(dropDownInfo.AllowDrop), dropDownInfo.AllowDrop));
            jObj.Add(new JProperty(nameof(dropDownInfo.AllowResize), dropDownInfo.AllowResize));
            jObj.Add(new JProperty(nameof(dropDownInfo.AutoDropDown), dropDownInfo.AutoDropDown));
            jObj.Add(new JProperty(nameof(dropDownInfo.AutoHideTouchKeyboard), dropDownInfo.AutoHideTouchKeyboard));
            jObj.Add(new JProperty(nameof(dropDownInfo.ClosingAnimation), dropDownInfo.ClosingAnimation));
            jObj.Add(new JProperty(nameof(dropDownInfo.Direction), dropDownInfo.Direction));
            jObj.Add(new JProperty(nameof(dropDownInfo.OpeningAnimation), dropDownInfo.OpeningAnimation));
            jObj.Add(new JProperty(nameof(dropDownInfo.ShowShadow), dropDownInfo.ShowShadow));
            jObj.Add(new JProperty(nameof(dropDownInfo.Size), $"{dropDownInfo.Size.Width}, {dropDownInfo.Size.Height}"));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="dropDownInfo">DropDownInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(DropDownInfo dropDownInfo, JToken prop)
        {
            var allowDrop = prop.SelectToken(nameof(DropDownInfo.AllowDrop));
            if (allowDrop != null)
            {
                dropDownInfo.AllowDrop = allowDrop.ToObject<bool>();
            }

            var allowResize = prop.SelectToken(nameof(DropDownInfo.AllowResize));
            if (allowResize != null)
            {
                dropDownInfo.AllowResize = allowResize.ToObject<bool>();
            }

            var autoDropDown = prop.SelectToken(nameof(DropDownInfo.AutoDropDown));
            if (autoDropDown != null)
            {
                dropDownInfo.AutoDropDown = autoDropDown.ToObject<bool>();
            }

            var autoHideTouchKeyboard = prop.SelectToken(nameof(DropDownInfo.AutoHideTouchKeyboard));
            if (autoHideTouchKeyboard != null)
            {
                dropDownInfo.AutoHideTouchKeyboard = autoHideTouchKeyboard.ToObject<AutoHideTouchKeyboard>();
            }

            var closingAnimation = prop.SelectToken(nameof(DropDownInfo.ClosingAnimation));
            if (closingAnimation != null)
            {
                dropDownInfo.ClosingAnimation = closingAnimation.ToObject<DropDownAnimation>();
            }

            var direction = prop.SelectToken(nameof(DropDownInfo.Direction));
            if (direction != null)
            {
                dropDownInfo.Direction = direction.ToObject<DropDownDirection>();
            }

            var openingAnimation = prop.SelectToken(nameof(DropDownInfo.OpeningAnimation));
            if (openingAnimation != null)
            {
                dropDownInfo.OpeningAnimation = openingAnimation.ToObject<DropDownAnimation>();
            }

            var showShadow = prop.SelectToken(nameof(DropDownInfo.ShowShadow));
            if (showShadow != null)
            {
                dropDownInfo.ShowShadow = showShadow.ToObject<bool>();
            }

            var size = prop.SelectToken(nameof(DropDownInfo.Size));
            if (size != null)
            {
                dropDownInfo.Size = size.ToObject<Size>();
            }
        }
    }
}
