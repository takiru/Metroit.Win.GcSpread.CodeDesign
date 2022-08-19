using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// GcDateTimeCellType のコンバーターを提供します。
    /// </summary>
    internal class GcDateTimeCellTypeConverter : FieldsEditorControlCellTypeConverter
    {
        /// <summary>
        /// 新しい GcDateTimeCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public GcDateTimeCellTypeConverter(GcDateTimeCellType cellType) : base(cellType)
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

            var c = CellType as GcDateTimeCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.AcceptsCrLf))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.AcceptsCrLf), c.AcceptsCrLf));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.AlternateText))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.AlternateText), DateAlternateTextInfoConverter.Serialize(c.AlternateText)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.DefaultActiveField))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.DefaultActiveField), c.SerializationDefaultActiveFieldIndex));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.DisplayFields))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.DisplayFields), DateTimeDisplayFieldCollectionInfoConverter.Serialize(c.DisplayFields)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.DropDown))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.DropDown), DropDownInfoConverter.Serialize(c.DropDown)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.DropDownCalendar))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.DropDownCalendar), DropDownCalendarInfoConverter.Serialize(c.DropDownCalendar)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.Fields))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.Fields), DateTimeFieldCollectionInfoConverter.Serialize(c.Fields)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.FieldsEditMode))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.FieldsEditMode), c.FieldsEditMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.FocusPosition))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.FocusPosition), c.FocusPosition));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.MaxDate))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.MaxDate), c.MaxDate));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.MaxMinBehavior))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.MaxMinBehavior), c.MaxMinBehavior));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.MinDate))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.MinDate), c.MinDate));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.PaintByControl))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.PaintByControl), c.PaintByControl));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.RecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.RecommendedValue), c.RecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.ShowRecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.ShowRecommendedValue), c.ShowRecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.SideButtons))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.SideButtons), SideButtonCollectionInfoConverter.Serialize(c.SideButtons)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.Spin))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.Spin), DateSpinConverter.Serialize(c.Spin)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcDateTimeCellType.ValidateMode))))
            {
                propsObj.Add(new JProperty(nameof(GcDateTimeCellType.ValidateMode), c.ValidateMode));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as GcDateTimeCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.AcceptsCrLf), true) == 0)
            {
                c.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.AlternateText), true) == 0)
            {
                DateAlternateTextInfoConverter.Deserialize(c.AlternateText, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DefaultActiveField), true) == 0)
            {
                c.SerializationDefaultActiveFieldIndex = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DisplayFields), true) == 0)
            {
                c.DisplayFields = DateTimeDisplayFieldCollectionInfoConverter.Deserialize((JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DropDown), true) == 0)
            {
                DropDownInfoConverter.Deserialize(c.DropDown, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DropDownCalendar), true) == 0)
            {
                DropDownCalendarInfoConverter.Deserialize(c.DropDownCalendar, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.Fields), true) == 0)
            {
                c.Fields = DateTimeFieldCollectionInfoConverter.Deserialize((JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.FieldsEditMode), true) == 0)
            {
                c.FieldsEditMode = prop.Value.ToObject<FieldsEditMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.FocusPosition), true) == 0)
            {
                c.FocusPosition = prop.Value.ToObject<FieldsEditorFocusCursorPosition>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.MaxDate), true) == 0)
            {
                c.MaxDate = prop.Value.ToObject<DateTime>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.MaxMinBehavior), true) == 0)
            {
                c.MaxMinBehavior = prop.Value.ToObject<MaxMinBehavior>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.MinDate), true) == 0)
            {
                c.MinDate = prop.Value.ToObject<DateTime>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.PaintByControl), true) == 0)
            {
                c.PaintByControl = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.RecommendedValue), true) == 0)
            {
                c.RecommendedValue = prop.Value.ToObject<DateTime?>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ShowRecommendedValue), true) == 0)
            {
                c.ShowRecommendedValue = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.SideButtons), true) == 0)
            {
                c.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.Spin), true) == 0)
            {
                DateSpinConverter.Deserialize(c.Spin, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ValidateMode), true) == 0)
            {
                c.ValidateMode = prop.Value.ToObject<ValidateModeEx>();
                return;
            }
        }
    }
}
