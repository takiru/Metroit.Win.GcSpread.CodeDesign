using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// ListGridLinesInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class ListGridLinesInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="listGridLinesInfo">ListGridLinesInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(ListGridLinesInfo listGridLinesInfo)
        {
            if (listGridLinesInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(ListGridLinesInfo.HorizontalLines), LineConverter.Serialize(listGridLinesInfo.HorizontalLines)));
            jObj.Add(new JProperty(nameof(ListGridLinesInfo.VerticalLines), LineConverter.Serialize(listGridLinesInfo.VerticalLines)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>ListGridLinesInfo オブジェクト。</returns>
        public static ListGridLinesInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new ListGridLinesInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする ListGridLinesInfo オブジェクト。</param>
        /// <returns>ListGridLinesInfo オブジェクト。</returns>
        public static ListGridLinesInfo Deserialize(JToken prop, ListGridLinesInfo source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new ListGridLinesInfo();
            if (source != null)
            {
                result.HorizontalLines = source.HorizontalLines;
                result.VerticalLines = source.VerticalLines;
            }

            var horizontalLines = prop.SelectToken(nameof(ListGridLinesInfo.HorizontalLines));
            if (horizontalLines != null)
            {
                result.HorizontalLines = LineConverter.Deserialize(horizontalLines, result.HorizontalLines);
            }

            var verticalLines = prop.SelectToken(nameof(ListGridLinesInfo.VerticalLines));
            if (verticalLines != null)
            {
                result.VerticalLines = LineConverter.Deserialize(verticalLines, result.VerticalLines);
            }

            return result;
        }
    }
}
