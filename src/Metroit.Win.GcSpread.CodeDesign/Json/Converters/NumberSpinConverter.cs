using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// NumberSpin オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class NumberSpinConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="numberSpin">NumberSpin オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(NumberSpin numberSpin)
        {
            if (numberSpin == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(NumberSpin.AllowSpin), numberSpin.AllowSpin));
            jObj.Add(new JProperty(nameof(NumberSpin.Increment), numberSpin.Increment));
            jObj.Add(new JProperty(nameof(NumberSpin.IncrementValue), numberSpin.IncrementValue));
            jObj.Add(new JProperty(nameof(NumberSpin.SpinMode), numberSpin.SpinMode));
            jObj.Add(new JProperty(nameof(NumberSpin.SpinOnKeys), numberSpin.SpinOnKeys));
            jObj.Add(new JProperty(nameof(NumberSpin.SpinOnWheel), numberSpin.SpinOnWheel));
            jObj.Add(new JProperty(nameof(NumberSpin.Wrap), numberSpin.Wrap));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="numberSpin">NumberSpin オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(NumberSpin numberSpin, JToken prop)
        {
            var allowSpin = prop.SelectToken(nameof(NumberSpin.AllowSpin));
            if (allowSpin != null)
            {
                numberSpin.AllowSpin = allowSpin.ToObject<bool>();
            }

            var increment = prop.SelectToken(nameof(NumberSpin.Increment));
            if (increment != null)
            {
                numberSpin.Increment = increment.ToObject<int>();
            }

            var incrementValue = prop.SelectToken(nameof(NumberSpin.IncrementValue));
            if (incrementValue != null)
            {
                numberSpin.IncrementValue = incrementValue.ToObject<decimal>();
            }

            var spinMode = prop.SelectToken(nameof(NumberSpin.SpinMode));
            if (spinMode != null)
            {
                numberSpin.SpinMode = spinMode.ToObject<NumberSpinMode>();
            }

            var spinOnKeys = prop.SelectToken(nameof(NumberSpin.SpinOnKeys));
            if (spinOnKeys != null)
            {
                numberSpin.SpinOnKeys = spinOnKeys.ToObject<bool>();
            }

            var spinOnWheel = prop.SelectToken(nameof(NumberSpin.SpinOnWheel));
            if (spinOnWheel != null)
            {
                numberSpin.SpinOnWheel = spinOnWheel.ToObject<bool>();
            }

            var wrap = prop.SelectToken(nameof(NumberSpin.Wrap));
            if (wrap != null)
            {
                numberSpin.Wrap = wrap.ToObject<bool>();
            }
        }
    }
}
