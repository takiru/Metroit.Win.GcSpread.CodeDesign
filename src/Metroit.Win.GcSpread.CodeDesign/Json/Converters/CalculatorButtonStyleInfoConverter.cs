using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System.Drawing;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// CalculatorButtonStyleInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class CalculatorButtonStyleInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="calculatorButtonStyleInfo">CalculatorButtonStyleInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(CalculatorButtonStyleInfo calculatorButtonStyleInfo)
        {
            if (calculatorButtonStyleInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(CalculatorButtonStyleInfo.BackColor), ColorTranslator.ToHtml(calculatorButtonStyleInfo.BackColor)));
            jObj.Add(new JProperty(nameof(CalculatorButtonStyleInfo.ForeColor), ColorTranslator.ToHtml(calculatorButtonStyleInfo.ForeColor)));
            jObj.Add(new JProperty(nameof(CalculatorButtonStyleInfo.TextEffect), calculatorButtonStyleInfo.TextEffect));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>CalculatorButtonStyleInfo オブジェクト。</returns>
        public static CalculatorButtonStyleInfo Deserialize(JToken prop)
        {
            return Deserialize(prop, new CalculatorButtonStyleInfo());
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <param name="source">ベースにする CalculatorButtonStyleInfo オブジェクト。</param>
        /// <returns>CalculatorButtonStyleInfo オブジェクト。</returns>
        public static CalculatorButtonStyleInfo Deserialize(JToken prop, CalculatorButtonStyleInfo source)
        {
            var result = new CalculatorButtonStyleInfo();
            if (source == null)
            {
                result.BackColor = source.BackColor;
                result.ForeColor = source.ForeColor;
                result.TextEffect = source.TextEffect;
            }

            var backColor = prop.SelectToken(nameof(CalculatorButtonStyleInfo.BackColor));
            if (backColor != null)
            {
                result.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var foreColor = prop.SelectToken(nameof(CalculatorButtonStyleInfo.ForeColor));
            if (foreColor != null)
            {
                result.ForeColor = ColorTranslator.FromHtml(foreColor.ToObject<string>());
            }

            var textEffect = prop.SelectToken(nameof(CalculatorButtonStyleInfo.TextEffect));
            if (textEffect != null)
            {
                result.TextEffect = textEffect.ToObject<TextEffect>();
            }

            return result;
        }
    }
}
