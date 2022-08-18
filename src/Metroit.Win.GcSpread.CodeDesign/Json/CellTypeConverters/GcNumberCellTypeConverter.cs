using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// GcNumberCellType のコンバーターを提供します。
    /// </summary>
    internal class GcNumberCellTypeConverter : FieldsEditorControlCellTypeConverter
    {
        /// <summary>
        /// 新しい GcNumberCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public GcNumberCellTypeConverter(GcNumberCellType cellType) : base(cellType)
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

            var c = CellType as GcNumberCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.AcceptsCrLf))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.AcceptsCrLf), c.AcceptsCrLf));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.AcceptsDecimal))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.AcceptsDecimal), c.AcceptsDecimal));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.AllowDeleteToNull))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.AllowDeleteToNull), c.AllowDeleteToNull));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.AlternateText))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.AlternateText), NumberAlternateTextInfoConverter.Serialize(c.AlternateText)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.DisplayFields))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.DisplayFields), NumberDisplayFieldCollectionInfoConverter.Serialize(c.DisplayFields)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.DropDown))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.DropDown), DropDownInfoConverter.Serialize(c.DropDown)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.DropDownCalculator))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.DropDownCalculator), DropDownCalculatorInfoConverter.Serialize(c.DropDownCalculator)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.Fields))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.Fields), NumberFieldsInfoConverter.Serialize(c.Fields)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.FocusPosition))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.FocusPosition), c.FocusPosition));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.MaxMinBehavior))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.MaxMinBehavior), c.MaxMinBehavior));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.MaxValue))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.MaxValue), c.MaxValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.MinValue))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.MinValue), c.MinValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.NegativeColor))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.NegativeColor), ColorTranslator.ToHtml(c.NegativeColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.PaintByControl))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.PaintByControl), c.PaintByControl));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.PopUpCalculator))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.PopUpCalculator), PopUpCalculatorInfoConverter.Serialize(c.PopUpCalculator)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.RecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.RecommendedValue), c.RecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.RoundPattern))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.RoundPattern), c.RoundPattern));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.ShowRecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.ShowRecommendedValue), c.ShowRecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.SideButtons))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.SideButtons), SideButtonCollectionInfoConverter.Serialize(c.SideButtons)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.Spin))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.Spin), NumberSpinConverter.Serialize(c.Spin)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.UseNegativeColor))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.UseNegativeColor), c.UseNegativeColor));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcNumberCellType.ValueSign))))
            {
                propsObj.Add(new JProperty(nameof(GcNumberCellType.ValueSign), c.ValueSign));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as GcNumberCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(GcNumberCellType.AcceptsCrLf), true) == 0)
            {
                c.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.AcceptsDecimal), true) == 0)
            {
                c.AcceptsDecimal = prop.Value.ToObject<DecimalMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.AllowDeleteToNull), true) == 0)
            {
                c.AllowDeleteToNull = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.AlternateText), true) == 0)
            {
                NumberAlternateTextInfoConverter.Deserialize(c.AlternateText, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.DisplayFields), true) == 0)
            {
                c.DisplayFields = NumberDisplayFieldCollectionInfoConverter.Deserialize((JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.DropDown), true) == 0)
            {
                DropDownInfoConverter.Deserialize(c.DropDown, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.DropDownCalculator), true) == 0)
            {
                DropDownCalculatorInfoConverter.Deserialize(c.DropDownCalculator, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.Fields), true) == 0)
            {
                NumberFieldsInfoConverter.Deserialize(c.Fields, (JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.FocusPosition), true) == 0)
            {
                c.FocusPosition = prop.Value.ToObject<EditorBaseFocusCursorPosition>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.MaxMinBehavior), true) == 0)
            {
                c.MaxMinBehavior = prop.Value.ToObject<MaxMinBehavior>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.MaxValue), true) == 0)
            {
                c.MaxValue = prop.Value.ToObject<decimal>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.MinValue), true) == 0)
            {
                c.MinValue = prop.Value.ToObject<decimal>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.NegativeColor), true) == 0)
            {
                c.NegativeColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.PaintByControl), true) == 0)
            {
                c.PaintByControl = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.PopUpCalculator), true) == 0)
            {
                PopUpCalculatorInfoConverter.Deserialize(c.PopUpCalculator, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.RecommendedValue), true) == 0)
            {
                c.RecommendedValue = prop.Value.ToObject<decimal>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.RoundPattern), true) == 0)
            {
                c.RoundPattern = prop.Value.ToObject<RoundPattern>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.ShowRecommendedValue), true) == 0)
            {
                c.ShowRecommendedValue = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.SideButtons), true) == 0)
            {
                c.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.Spin), true) == 0)
            {
                NumberSpinConverter.Deserialize(c.Spin, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.UseNegativeColor), true) == 0)
            {
                c.UseNegativeColor = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcNumberCellType.ValueSign), true) == 0)
            {
                c.ValueSign = prop.Value.ToObject<ValueSignControl>();
                return;
            }
        }
    }
}
