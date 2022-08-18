using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// AlternateTextInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class AlternateTextInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="alternateTextInfo">AlternateTextInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(AlternateTextInfo alternateTextInfo)
        {
            if (alternateTextInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(AlternateTextInfo.ForeColor), ColorTranslator.ToHtml(alternateTextInfo.ForeColor)));
            jObj.Add(new JProperty(nameof(AlternateTextInfo.Text), alternateTextInfo.Text));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>AlternateTextInfo オブジェクト。</returns>
        public static AlternateTextInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new AlternateTextInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする AlternateTextInfo オブジェクト。</param>
        /// <returns>AlternateTextInfo オブジェクト。</returns>
        public static AlternateTextInfo Deserialize(JToken prop, AlternateTextInfo source)
        {
            if (prop.ToObject<object>() == null && source == null)
            {
                return null;
            }

            var result = new AlternateTextInfo();
            if (source != null)
            {
                result.ForeColor = source.ForeColor;
                result.Text = source.Text;
            }

            var foreColor = prop.SelectToken(nameof(AlternateTextInfo.ForeColor));
            if (foreColor != null)
            {
                result.ForeColor = ColorTranslator.FromHtml(foreColor.Value<string>());
            }

            var text = prop.SelectToken(nameof(AlternateTextInfo.Text));
            if (text != null)
            {
                result.Text = text.Value<string>();
            }

            return result;
        }
    }
}
