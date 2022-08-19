using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DateSpin オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DateSpinConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dateSpin">DateSpin オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DateSpin dateSpin)
        {
            if (dateSpin == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DateSpin.AllowSpin), dateSpin.AllowSpin));
            jObj.Add(new JProperty(nameof(DateSpin.Increment), dateSpin.Increment));
            jObj.Add(new JProperty(nameof(DateSpin.IncrementValue), dateSpin.IncrementValue));
            jObj.Add(new JProperty(nameof(DateSpin.SpinMode), dateSpin.SpinMode));
            jObj.Add(new JProperty(nameof(DateSpin.SpinOnKeys), dateSpin.SpinOnKeys));
            jObj.Add(new JProperty(nameof(DateSpin.SpinOnWheel), dateSpin.SpinOnWheel));
            jObj.Add(new JProperty(nameof(DateSpin.Wrap), dateSpin.Wrap));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="dateSpin">DateSpin オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(DateSpin dateSpin, JToken prop)
        {
            var allowSpin = prop.SelectToken(nameof(DateSpin.AllowSpin));
            if (allowSpin != null)
            {
                dateSpin.AllowSpin = allowSpin.ToObject<bool>();
            }

            var increment = prop.SelectToken(nameof(DateSpin.Increment));
            if (increment != null)
            {
                dateSpin.Increment = increment.ToObject<int>();
            }

            var incrementValue = prop.SelectToken(nameof(DateSpin.IncrementValue));
            if (incrementValue != null)
            {
                dateSpin.IncrementValue = incrementValue.ToObject<TimeSpan>();
            }

            var spinMode = prop.SelectToken(nameof(DateSpin.SpinMode));
            if (spinMode != null)
            {
                dateSpin.SpinMode = spinMode.ToObject<DateSpinMode>();
            }

            var spinOnKeys = prop.SelectToken(nameof(DateSpin.SpinOnKeys));
            if (spinOnKeys != null)
            {
                dateSpin.SpinOnKeys = spinOnKeys.ToObject<bool>();
            }

            var spinOnWheel = prop.SelectToken(nameof(DateSpin.SpinOnWheel));
            if (spinOnWheel != null)
            {
                dateSpin.SpinOnWheel = spinOnWheel.ToObject<bool>();
            }

            var wrap = prop.SelectToken(nameof(DateSpin.Wrap));
            if (wrap != null)
            {
                dateSpin.Wrap = wrap.ToObject<bool>();
            }
        }
    }
}
