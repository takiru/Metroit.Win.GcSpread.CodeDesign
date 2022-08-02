using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// StatusBarInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class StatusBarInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="statusBarInfo">StatusBarInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(StatusBarInfo statusBarInfo)
        {
            if (statusBarInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(StatusBarInfo.Text), statusBarInfo.Text));
            jObj.Add(new JProperty(nameof(StatusBarInfo.Visible), statusBarInfo.Visible));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="statusBarInfo">StatusBarInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(StatusBarInfo statusBarInfo, JToken prop)
        {
            var text = prop.SelectToken(nameof(StatusBarInfo.Text));
            if (text != null)
            {
                statusBarInfo.Text = text.ToObject<string>();
            }

            var visible = prop.SelectToken(nameof(StatusBarInfo.Visible));
            if (visible != null)
            {
                statusBarInfo.Visible = visible.ToObject<bool>();
            }
        }
    }
}
