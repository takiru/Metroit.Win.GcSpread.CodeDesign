using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// Spin オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class SpinConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="spin">Spin オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(Spin spin)
        {
            if (spin == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(Spin.AllowSpin), spin.AllowSpin));
            jObj.Add(new JProperty(nameof(Spin.Increment), spin.Increment));
            jObj.Add(new JProperty(nameof(Spin.SpinOnKeys), spin.SpinOnKeys));
            jObj.Add(new JProperty(nameof(Spin.SpinOnWheel), spin.SpinOnWheel));
            jObj.Add(new JProperty(nameof(Spin.Wrap), spin.Wrap));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="spin">Spin オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(Spin spin, JToken prop)
        {
            var allowSpin = prop.SelectToken(nameof(Spin.AllowSpin));
            if (allowSpin != null)
            {
                spin.AllowSpin = allowSpin.ToObject<bool>();
            }

            var increment = prop.SelectToken(nameof(Spin.Increment));
            if (increment != null)
            {
                spin.Increment = increment.ToObject<int>();
            }

            var spinOnKeys = prop.SelectToken(nameof(Spin.SpinOnKeys));
            if (spinOnKeys != null)
            {
                spin.SpinOnKeys = spinOnKeys.ToObject<bool>();
            }

            var spinOnWheel = prop.SelectToken(nameof(Spin.SpinOnWheel));
            if (spinOnWheel != null)
            {
                spin.SpinOnWheel = spinOnWheel.ToObject<bool>();
            }

            var wrap = prop.SelectToken(nameof(Spin.Wrap));
            if (wrap != null)
            {
                spin.Wrap = wrap.ToObject<bool>();
            }
        }
    }
}
