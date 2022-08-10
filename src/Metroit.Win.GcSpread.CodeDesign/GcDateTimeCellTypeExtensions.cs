using FarPoint.Win.SuperEdit;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System;

namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// GcDateTimeCellType のJSONシリアライズ、デシリアライズを提供します。
    /// </summary>
    public static class GcDateTimeCellTypeExtensions
    {
        /// <summary>
        /// GcDateTimeCellType をJSONシリアライズします。
        /// SubEditor, TouchToolBar はシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcDateTimeCellType オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static string SerializeJson(this GcDateTimeCellType cellType)
        {
            var result = new JObject();
            result.Add(new JProperty(nameof(cellType.AcceptsArrowKeys), cellType.AcceptsArrowKeys));
            result.Add(new JProperty(nameof(cellType.AcceptsCrLf), cellType.AcceptsCrLf));
            result.Add(new JProperty(nameof(cellType.AlternateText), DateAlternateTextInfoConverter.Serialize(cellType.AlternateText)));
            result.Add(new JProperty(nameof(cellType.BackgroundImage), PictureConverter.Serialize(cellType.BackgroundImage)));
            result.Add(new JProperty(nameof(cellType.ClipContent), cellType.ClipContent));
            result.Add(new JProperty(nameof(cellType.DefaultActiveField), cellType.SerializationDefaultActiveFieldIndex));
            result.Add(new JProperty(nameof(cellType.DisplayFields), DateTimeDisplayFieldCollectionInfoConverter.Serialize(cellType.DisplayFields)));
            result.Add(new JProperty(nameof(cellType.DropDown), DateTimeDropDownInfoConverter.Serialize(cellType.DropDown)));
            result.Add(new JProperty(nameof(cellType.DropDownCalendar), DropDownCalendarInfoConverter.Serialize(cellType.DropDownCalendar)));
            result.Add(new JProperty(nameof(cellType.DropDownPicker), DateTimeDropDownPickerInfoConverter.Serialize(cellType.DropDownPicker)));
            result.Add(new JProperty(nameof(cellType.EditMode), cellType.EditMode));
            result.Add(new JProperty(nameof(cellType.ExcelExportFormat), cellType.ExcelExportFormat));
            result.Add(new JProperty(nameof(cellType.ExitOnLastChar), cellType.ExitOnLastChar));
            result.Add(new JProperty(nameof(cellType.Fields), DateTimeFieldCollectionInfoConverter.Serialize(cellType.Fields)));
            result.Add(new JProperty(nameof(cellType.FieldsEditMode), cellType.FieldsEditMode));
            result.Add(new JProperty(nameof(cellType.FocusPosition), cellType.FocusPosition));
            result.Add(new JProperty(nameof(cellType.MaxDate), cellType.MaxDate));
            result.Add(new JProperty(nameof(cellType.MaxMinBehavior), cellType.MaxMinBehavior));
            result.Add(new JProperty(nameof(cellType.MinDate), cellType.MinDate));
            result.Add(new JProperty(nameof(cellType.PaintByControl), cellType.PaintByControl));
            result.Add(new JProperty(nameof(cellType.PromptChar), cellType.PromptChar.ToString()));
            result.Add(new JProperty(nameof(cellType.ReadOnly), cellType.ReadOnly));
            result.Add(new JProperty(nameof(cellType.RecommendedValue), cellType.RecommendedValue));
            result.Add(new JProperty(nameof(cellType.ShortcutKeys), ShortcutDictionaryConverter.Serialize(cellType.ShortcutKeys)));
            result.Add(new JProperty(nameof(cellType.ShowLiterals), cellType.ShowLiterals));
            result.Add(new JProperty(nameof(cellType.ShowRecommendedValue), cellType.ShowRecommendedValue));
            result.Add(new JProperty(nameof(cellType.ShowTouchToolBar), cellType.ShowTouchToolBar));
            result.Add(new JProperty(nameof(cellType.SideButtons), SideButtonCollectionInfoConverter.Serialize(cellType.SideButtons)));
            result.Add(new JProperty(nameof(cellType.Spin), DateSpinConverter.Serialize(cellType.Spin)));
            result.Add(new JProperty(nameof(cellType.Static), cellType.Static));
            result.Add(new JProperty(nameof(cellType.TabAction), cellType.TabAction));
            result.Add(new JProperty(nameof(cellType.TouchContextMenuScale), cellType.TouchContextMenuScale));
            result.Add(new JProperty(nameof(cellType.UseSpreadDropDownButtonRender), cellType.UseSpreadDropDownButtonRender));
            result.Add(new JProperty(nameof(cellType.ValidateMode), cellType.ValidateMode));
            
            return Newtonsoft.Json.JsonConvert.SerializeObject(result);
        }

        /// <summary>
        /// GcDateTimeCellType をJSONデシリアライズします。
        /// SubEditor, TouchToolBar はデシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcDateTimeCellType オブジェクト。</param>
        /// <param name="cellTypeProps">CellType のプロパティのJSON文字列。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static void DeserializeJson(this GcDateTimeCellType cellType, string cellTypeProps)
        {
            if (string.IsNullOrEmpty(cellTypeProps))
            {
                return;
            }

            var props = JObject.Parse(cellTypeProps);
            foreach (var prop in props)
            {
                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.AcceptsArrowKeys), true) == 0)
                {
                    cellType.AcceptsArrowKeys = prop.Value.ToObject<AcceptsArrowKeys>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.AcceptsCrLf), true) == 0)
                {
                    cellType.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.AlternateText), true) == 0)
                {
                    DateAlternateTextInfoConverter.Deserialize(cellType.AlternateText, prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.BackgroundImage), true) == 0)
                {
                    cellType.BackgroundImage = PictureConverter.Deserialize(prop.Value, cellType.BackgroundImage);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ClipContent), true) == 0)
                {
                    cellType.ClipContent = prop.Value.ToObject<ClipContent>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DefaultActiveField), true) == 0)
                {
                    cellType.SerializationDefaultActiveFieldIndex = prop.Value.ToObject<int>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DisplayFields), true) == 0)
                {
                    cellType.DisplayFields = DateTimeDisplayFieldCollectionInfoConverter.Deserialize((JArray)prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DropDown), true) == 0)
                {
                    DateTimeDropDownInfoConverter.Deserialize(cellType.DropDown, prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DropDownCalendar), true) == 0)
                {
                    DropDownCalendarInfoConverter.Deserialize(cellType.DropDownCalendar, prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.DropDownPicker), true) == 0)
                {
                    DateTimeDropDownPickerInfoConverter.Deserialize(cellType.DropDownPicker, prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.EditMode), true) == 0)
                {
                    cellType.EditMode = prop.Value.ToObject<EditMode>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ExcelExportFormat), true) == 0)
                {
                    cellType.ExcelExportFormat = prop.Value.ToObject<string>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ExitOnLastChar), true) == 0)
                {
                    cellType.ExitOnLastChar = prop.Value.ToObject<bool>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.Fields), true) == 0)
                {
                    cellType.Fields = DateTimeFieldCollectionInfoConverter.Deserialize((JArray)prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.FieldsEditMode), true) == 0)
                {
                    cellType.FieldsEditMode = prop.Value.ToObject<FieldsEditMode>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.FocusPosition), true) == 0)
                {
                    cellType.FocusPosition = prop.Value.ToObject<FieldsEditorFocusCursorPosition>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.MaxDate), true) == 0)
                {
                    cellType.MaxDate = prop.Value.ToObject<DateTime>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.MaxMinBehavior), true) == 0)
                {
                    cellType.MaxMinBehavior = prop.Value.ToObject<MaxMinBehavior>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.MinDate), true) == 0)
                {
                    cellType.MinDate = prop.Value.ToObject<DateTime>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.PaintByControl), true) == 0)
                {
                    cellType.PaintByControl = prop.Value.ToObject<bool>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.PromptChar), true) == 0)
                {
                    cellType.PromptChar = prop.Value.ToObject<char>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ReadOnly), true) == 0)
                {
                    cellType.ReadOnly = prop.Value.ToObject<bool>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.RecommendedValue), true) == 0)
                {
                    cellType.RecommendedValue = prop.Value.ToObject<DateTime?>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ShortcutKeys), true) == 0)
                {
                    cellType.ShortcutKeys.Clear();
                    ShortcutDictionaryConverter.Deserialize(cellType.ShortcutKeys, prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ShowLiterals), true) == 0)
                {
                    cellType.ShowLiterals = prop.Value.ToObject<ShowLiterals>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ShowRecommendedValue), true) == 0)
                {
                    cellType.ShowRecommendedValue = prop.Value.ToObject<bool>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ShowTouchToolBar), true) == 0)
                {
                    cellType.ShowTouchToolBar = prop.Value.ToObject<TouchToolBarDisplayOptions>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.SideButtons), true) == 0)
                {
                    cellType.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.Spin), true) == 0)
                {
                    DateSpinConverter.Deserialize(cellType.Spin, prop.Value);
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.Static), true) == 0)
                {
                    cellType.Static = prop.Value.ToObject<bool>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.TabAction), true) == 0)
                {
                    cellType.TabAction = prop.Value.ToObject<TabAction>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.TouchContextMenuScale), true) == 0)
                {
                    cellType.TouchContextMenuScale = prop.Value.ToObject<float>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.UseSpreadDropDownButtonRender), true) == 0)
                {
                    cellType.UseSpreadDropDownButtonRender = prop.Value.ToObject<bool>();
                    continue;
                }

                if (string.Compare(prop.Key, nameof(GcDateTimeCellType.ValidateMode), true) == 0)
                {
                    cellType.ValidateMode = prop.Value.ToObject<ValidateModeEx>();
                    continue;
                }
            }
        }
    }
}
