using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// NumberCellType のコンバーターを提供します。
    /// </summary>
    internal class NumberCellTypeConverter : EditBaseCellTypeConverter
    {
        /// <summary>
        /// 新しい NumberCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public NumberCellTypeConverter(NumberCellType cellType) : base(cellType)
        {

        }

        /// <summary>
        /// セルタイプのプロパティをシリアライズします。
        /// </summary>
        /// <param name="propsObj">プロパティオブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        protected override void SerializeProp(JObject propsObj, string[] includeProps)
        {
            base.SerializeProp(propsObj, includeProps);

            var c = CellType as NumberCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.DecimalPlaces))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.DecimalPlaces), c.DecimalPlaces));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.DecimalSeparator))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.DecimalSeparator), c.DecimalSeparator));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.FixedPoint))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.FixedPoint), c.FixedPoint));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.FractionConvertWholeNumbers))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.FractionConvertWholeNumbers), c.FractionConvertWholeNumbers));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.FractionCustomFormat))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.FractionCustomFormat), c.FractionCustomFormat));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.FractionDenominatorDigits))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.FractionDenominatorDigits), c.FractionDenominatorDigits));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.FractionDenominatorPrecision))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.FractionDenominatorPrecision), c.FractionDenominatorPrecision));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.FractionMode))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.FractionMode), c.FractionMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.FractionRenderOnly))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.FractionRenderOnly), c.FractionRenderOnly));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.LeadingZero))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.LeadingZero), c.LeadingZero));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.MaximumValue))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.MaximumValue), c.MaximumValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.MinimumValue))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.MinimumValue), c.MinimumValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.NegativeFormat))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.NegativeFormat), c.NegativeFormat));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.NegativeRed))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.NegativeRed), c.NegativeRed));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.OverflowCharacter))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.OverflowCharacter), c.OverflowCharacter.ToString()));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.Separator))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.Separator), c.Separator));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.ShowSeparator))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.ShowSeparator), c.ShowSeparator));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.SpinButton))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.SpinButton), c.SpinButton));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.SpinDecimalIncrement))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.SpinDecimalIncrement), c.SpinDecimalIncrement));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.SpinIntegerIncrement))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.SpinIntegerIncrement), c.SpinIntegerIncrement));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(NumberCellType.SpinWrap))))
            {
                propsObj.Add(new JProperty(nameof(NumberCellType.SpinWrap), c.SpinWrap));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as NumberCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(NumberCellType.DecimalPlaces), true) == 0)
            {
                c.DecimalPlaces = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.DecimalSeparator), true) == 0)
            {
                c.DecimalSeparator = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.FixedPoint), true) == 0)
            {
                c.FixedPoint = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.FractionConvertWholeNumbers), true) == 0)
            {
                c.FractionConvertWholeNumbers = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.FractionCustomFormat), true) == 0)
            {
                c.FractionCustomFormat = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.FractionDenominatorDigits), true) == 0)
            {
                c.FractionDenominatorDigits = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.FractionDenominatorPrecision), true) == 0)
            {
                c.FractionDenominatorPrecision = prop.Value.ToObject<FractionDenominatorPrecision>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.FractionMode), true) == 0)
            {
                c.FractionMode = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.FractionRenderOnly), true) == 0)
            {
                c.FractionRenderOnly = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.LeadingZero), true) == 0)
            {
                c.LeadingZero = prop.Value.ToObject<LeadingZero>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.MaximumValue), true) == 0)
            {
                c.MaximumValue = prop.Value.ToObject<double>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.MinimumValue), true) == 0)
            {
                c.MinimumValue = prop.Value.ToObject<double>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.NegativeFormat), true) == 0)
            {
                c.NegativeFormat = prop.Value.ToObject<NegativeFormat>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.NegativeRed), true) == 0)
            {
                c.NegativeRed = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.OverflowCharacter), true) == 0)
            {
                c.OverflowCharacter = prop.Value.ToObject<char>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.Separator), true) == 0)
            {
                c.Separator = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.ShowSeparator), true) == 0)
            {
                c.ShowSeparator = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.SpinButton), true) == 0)
            {
                c.SpinButton = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.SpinDecimalIncrement), true) == 0)
            {
                c.SpinDecimalIncrement = prop.Value.ToObject<float>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.SpinIntegerIncrement), true) == 0)
            {
                c.SpinIntegerIncrement = prop.Value.ToObject<float>();
                return;
            }
            if (string.Compare(prop.Key, nameof(NumberCellType.SpinWrap), true) == 0)
            {
                c.SpinWrap = prop.Value.ToObject<bool>();
                return;
            }
        }
    }
}
