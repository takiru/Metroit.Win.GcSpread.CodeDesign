using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// Font オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class FontConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="font">Font オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(Font font)
        {
            if (font == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(Font.FontFamily.Name), font.FontFamily.Name));
            jObj.Add(new JProperty(nameof(Font.Size), font.Size));
            jObj.Add(new JProperty(nameof(Font.Style), font.Style));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>Font オブジェクト。</returns>
        public static Font Deserialize(JToken prop)
        {
            var name = prop.SelectToken(nameof(Font.Name));
            var size = prop.SelectToken(nameof(Font.Size));
            var style = prop.SelectToken(nameof(Font.Style));

            // NOTE: Font オブジェクトは、指定可能な情報の指定がすべてないと処理不可
            if (name != null && size != null)
            {
                return new Font(name.ToObject<string>(), size.ToObject<float>());
            }
            if (name != null && size != null && style != null)
            {
                return new Font(name.ToObject<string>(), size.ToObject<float>(), style.ToObject<FontStyle>());
            }
            
            return null;
        }
    }
}
