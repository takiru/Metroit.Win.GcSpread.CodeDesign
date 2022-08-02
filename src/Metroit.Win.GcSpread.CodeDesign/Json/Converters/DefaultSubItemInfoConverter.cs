using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DefaultSubItemInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DefaultSubItemInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="defaultSubItemInfo">DefaultSubItemInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DefaultSubItemInfo defaultSubItemInfo)
        {
            if (defaultSubItemInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DefaultSubItemInfo.ContentAlignment), defaultSubItemInfo.ContentAlignment));
            jObj.Add(new JProperty(nameof(DefaultSubItemInfo.Ellipsis), defaultSubItemInfo.Ellipsis));
            jObj.Add(new JProperty(nameof(DefaultSubItemInfo.Padding),
                $"{defaultSubItemInfo.Padding.Left}, {defaultSubItemInfo.Padding.Top}, {defaultSubItemInfo.Padding.Right}, {defaultSubItemInfo.Padding.Bottom}"));
            jObj.Add(new JProperty(nameof(DefaultSubItemInfo.WordWrap), defaultSubItemInfo.WordWrap));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>DefaultSubItemInfo オブジェクト。</returns>
        public static DefaultSubItemInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new DefaultSubItemInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする DefaultSubItemInfo オブジェクト。</param>
        /// <returns>DefaultSubItemInfo オブジェクト。</returns>
        public static DefaultSubItemInfo Deserialize(JToken prop, DefaultSubItemInfo source)
        {
            var result = new DefaultSubItemInfo();
            if (source != null)
            {
                result.ContentAlignment = source.ContentAlignment;
                result.Ellipsis = source.Ellipsis;
                result.Padding = source.Padding;
                result.WordWrap = source.WordWrap;
            }

            var contentAlignment = prop.SelectToken(nameof(DefaultSubItemInfo.ContentAlignment));
            if (contentAlignment != null)
            {
                result.ContentAlignment = contentAlignment.ToObject<ContentAlignment>();
            }

            var ellipsis = prop.SelectToken(nameof(DefaultSubItemInfo.Ellipsis));
            if (ellipsis != null)
            {
                result.Ellipsis = ellipsis.ToObject<bool>();
            }

            var padding = prop.SelectToken(nameof(DefaultSubItemInfo.Padding));
            if (padding != null)
            {
                result.Padding = padding.ToObject<Padding>();
            }

            var wordWrap = prop.SelectToken(nameof(DefaultSubItemInfo.WordWrap));
            if (wordWrap != null)
            {
                result.WordWrap = wordWrap.ToObject<bool>();
            }

            return result;
        }
    }
}
