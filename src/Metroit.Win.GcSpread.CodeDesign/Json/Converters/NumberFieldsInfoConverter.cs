using GrapeCity.Win.Spread.InputMan.CellType;
using GrapeCity.Win.Spread.InputMan.CellType.Fields;
using Newtonsoft.Json.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// NumberFieldsInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal class NumberFieldsInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="numberFieldInfos">NumberFieldsInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JArray Serialize(NumberFieldsInfo numberFieldInfos)
        {
            if (numberFieldInfos == null)
            {
                return null;
            }

            var result = new JArray();

            foreach (NumberFieldInfo field in numberFieldInfos)
            {
                var jobj = JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(field));
                jobj.AddFirst(new JProperty("Type", field.GetType().Name));
                result.Add(jobj);
            }

            return result;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        /// <returns>NumberFieldsInfo オブジェクト。</returns>
        public static void Deserialize(NumberFieldsInfo numberFieldInfos, JArray props)
        {
            foreach (JToken prop in props)
            {
                var typeName = prop.SelectToken("Type").ToObject<string>();
                if (string.Compare(typeName, nameof(NumberSignFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberSignFieldInfo>();
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.BackColor)) != null) {
                        numberFieldInfos.SignPrefix.BackColor = info.BackColor;
                    }
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.Font)) != null)
                    {
                        numberFieldInfos.SignPrefix.Font = info.Font;
                    }
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.ForeColor)) != null)
                    {
                        numberFieldInfos.SignPrefix.ForeColor = info.ForeColor;
                    }
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.Margin)) != null)
                    {
                        numberFieldInfos.SignPrefix.Margin = info.Margin;
                    }
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.NegativePattern)) != null)
                    {
                        numberFieldInfos.SignPrefix.NegativePattern = info.NegativePattern;
                    }
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.Padding)) != null)
                    {
                        numberFieldInfos.SignPrefix.Padding = info.Padding;
                    }
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.PositivePattern)) != null)
                    {
                        numberFieldInfos.SignPrefix.PositivePattern = info.PositivePattern;
                    }
                    if (prop.SelectToken(nameof(NumberSignFieldInfo.Text)) != null)
                    {
                        numberFieldInfos.SignPrefix.Text = info.Text;
                    }
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberIntegerPartFieldInfo), true) == 0)
                {
                    var info = prop.ToObject<NumberIntegerPartFieldInfo>();
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.BackColor)) != null)
                    {
                        numberFieldInfos.IntegerPart.BackColor = info.BackColor;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.Font)) != null)
                    {
                        numberFieldInfos.IntegerPart.Font = info.Font;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.ForeColor)) != null)
                    {
                        numberFieldInfos.IntegerPart.ForeColor = info.ForeColor;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.GroupSeparator)) != null)
                    {
                        numberFieldInfos.IntegerPart.GroupSeparator = info.GroupSeparator;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.GroupSizes)) != null)
                    {
                        numberFieldInfos.IntegerPart.GroupSizes = info.GroupSizes;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.Margin)) != null)
                    {
                        numberFieldInfos.IntegerPart.Margin = info.Margin;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.MaxDigits)) != null)
                    {
                        numberFieldInfos.IntegerPart.MaxDigits = info.MaxDigits;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.MinDigits)) != null)
                    {
                        numberFieldInfos.IntegerPart.MinDigits = info.MinDigits;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.Padding)) != null)
                    {
                        numberFieldInfos.IntegerPart.Padding = info.Padding;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.SpinIncrement)) != null)
                    {
                        numberFieldInfos.IntegerPart.SpinIncrement = info.SpinIncrement;
                    }
                    if (prop.SelectToken(nameof(NumberIntegerPartFieldInfo.Text)) != null)
                    {
                        numberFieldInfos.IntegerPart.Text = info.Text;
                    }
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberDecimalSeparatorFieldInfo), true) == 0)
                {
                    var info  = prop.ToObject<NumberDecimalSeparatorFieldInfo>();
                    if (prop.SelectToken(nameof(NumberDecimalSeparatorFieldInfo.BackColor)) != null)
                    {
                        numberFieldInfos.DecimalSeparator.BackColor = info.BackColor;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalSeparatorFieldInfo.DecimalSeparator)) != null)
                    {
                        numberFieldInfos.DecimalSeparator.DecimalSeparator = info.DecimalSeparator;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalSeparatorFieldInfo.Font)) != null)
                    {
                        numberFieldInfos.DecimalSeparator.Font = info.Font;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalSeparatorFieldInfo.ForeColor)) != null)
                    {
                        numberFieldInfos.DecimalSeparator.ForeColor = info.ForeColor;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalSeparatorFieldInfo.Margin)) != null)
                    {
                        numberFieldInfos.DecimalSeparator.Margin = info.Margin;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalSeparatorFieldInfo.Padding)) != null)
                    {
                        numberFieldInfos.DecimalSeparator.Padding = info.Padding;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalSeparatorFieldInfo.Text)) != null)
                    {
                        numberFieldInfos.DecimalSeparator.Text = info.Text;
                    }
                    continue;
                }
                if (string.Compare(typeName, nameof(NumberDecimalPartFieldInfo), true) == 0)
                {
                    var info  = prop.ToObject<NumberDecimalPartFieldInfo>();
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.BackColor)) != null)
                    {
                        numberFieldInfos.DecimalPart.BackColor = info.BackColor;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.Font)) != null)
                    {
                        numberFieldInfos.DecimalPart.Font = info.Font;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.ForeColor)) != null)
                    {
                        numberFieldInfos.DecimalPart.ForeColor = info.ForeColor;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.Margin)) != null)
                    {
                        numberFieldInfos.DecimalPart.Margin = info.Margin;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.MaxDigits)) != null)
                    {
                        numberFieldInfos.DecimalPart.MaxDigits = info.MaxDigits;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.MinDigits)) != null)
                    {
                        numberFieldInfos.DecimalPart.MinDigits = info.MinDigits;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.Padding)) != null)
                    {
                        numberFieldInfos.DecimalPart.Padding = info.Padding;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.SpinIncrement)) != null)
                    {
                        numberFieldInfos.DecimalPart.SpinIncrement = info.SpinIncrement;
                    }
                    if (prop.SelectToken(nameof(NumberDecimalPartFieldInfo.Text)) != null)
                    {
                        numberFieldInfos.DecimalPart.Text = info.Text;
                    }
                    continue;
                }
            }
        }
    }
}
