using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// Line オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class LineConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// LineWidth はシリアライズを行いません。
        /// </summary>
        /// <param name="line">Line オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         LineWidth: 読み取り専用でデシリアライズできないため。
        public static JObject Serialize(Line line)
        {
            if (line == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(Line.Color), ColorTranslator.ToHtml(line.Color)));
            jObj.Add(new JProperty(nameof(Line.Style), line.Style));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>Line オブジェクト。</returns>
        public static Line Deserialize(JToken prop)
        {
            return Deserialize(prop, new Line());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// LineWidth はデシリアライズを行いません。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする Line オブジェクト。</param>
        /// <returns>Line オブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         LineWidth: 読み取り専用でデシリアライズできないため。
        public static Line Deserialize(JToken prop, Line source)
        {
            var result = new Line();
            if (source != null)
            {
                result.Color = source.Color;
                result.Style = source.Style;
            }

            var color = prop.SelectToken(nameof(Line.Color));
            if (color != null)
            {
                result.Color = ColorTranslator.FromHtml(color.ToObject<string>());
            }

            var style = prop.SelectToken(nameof(Line.Style));
            if (style != null)
            {
                result.Style = style.ToObject<LineStyle>();
            }

            return result;
        }
    }
}
