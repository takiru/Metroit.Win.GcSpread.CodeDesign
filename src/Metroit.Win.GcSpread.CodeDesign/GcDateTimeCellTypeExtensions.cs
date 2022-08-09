using FarPoint.Win.SuperEdit;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Windows.Forms;

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
            result.Add(new JProperty(nameof(cellType.DefaultActiveField), cellType.DefaultActiveField));
            result.Add(new JProperty(nameof(cellType.DisplayFields), DateTimeDisplayFieldCollectionInfoConverter.Serialize(cellType.DisplayFields)));
            result.Add(new JProperty(nameof(cellType.DropDown), DateTimeDropDownInfoConverter.Serialize(cellType.DropDown)));
            result.Add(new JProperty(nameof(cellType.DropDownCalendar), cellType.DropDownCalendar));    // TODO
            result.Add(new JProperty(nameof(cellType.DropDownPicker), cellType.DropDownPicker));    // TODO
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
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AcceptsArrowKeys), true) == 0)
                {
                    cellType.AcceptsArrowKeys = prop.Value.ToObject<AcceptsArrowKeys>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AcceptsCrLf), true) == 0)
                {
                    cellType.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AcceptsTabChar), true) == 0)
                {
                    cellType.AcceptsTabChar = prop.Value.ToObject<TabCharMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AllowSpace), true) == 0)
                {
                    cellType.AllowSpace = prop.Value.ToObject<AllowSpace>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AlternateText), true) == 0)
                {
                    TextBoxAlternateTextInfoConverter.Deserialize(cellType.AlternateText, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoComplete), true) == 0)
                {
                    AutoCompleteInfoConverter.Deserialize(cellType.AutoComplete, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoCompleteCustomSource), true) == 0)
                {
                    cellType.AutoCompleteCustomSource = prop.Value.ToObject<AutoCompleteStringCollection>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoCompleteMode), true) == 0)
                {
                    cellType.AutoCompleteMode = prop.Value.ToObject<AutoCompleteMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoCompleteSource), true) == 0)
                {
                    cellType.AutoCompleteSource = prop.Value.ToObject<AutoCompleteSource>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.AutoConvert), true) == 0)
                {
                    cellType.AutoConvert = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.BackgroundImage), true) == 0)
                {
                    cellType.BackgroundImage = PictureConverter.Deserialize(prop.Value, cellType.BackgroundImage);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.DisplayAlignment), true) == 0)
                {
                    cellType.DisplayAlignment = prop.Value.ToObject<DisplayAlignment>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.DropDown), true) == 0)
                {
                    DropDownInfoConverter.Deserialize(cellType.DropDown, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.DropDownEditor), true) == 0)
                {
                    DropDownEditorInfoConverter.Deserialize(cellType.DropDownEditor, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.EditMode), true) == 0)
                {
                    cellType.EditMode = prop.Value.ToObject<EditMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.Ellipsis), true) == 0)
                {
                    cellType.Ellipsis = prop.Value.ToObject<EllipsisMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.EllipsisString), true) == 0)
                {
                    cellType.EllipsisString = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ExcelExportFormat), true) == 0)
                {
                    cellType.ExcelExportFormat = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ExitOnLastChar), true) == 0)
                {
                    cellType.ExitOnLastChar = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.FocusPosition), true) == 0)
                {
                    cellType.FocusPosition = prop.Value.ToObject<EditorBaseFocusCursorPosition>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.FormatString), true) == 0)
                {
                    cellType.FormatString = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.GridLine), true) == 0)
                {
                    cellType.GridLine = LineConverter.Deserialize(prop.Value, cellType.GridLine);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.LineSpace), true) == 0)
                {
                    cellType.LineSpace = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.MaxLength), true) == 0)
                {
                    cellType.MaxLength = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.MaxLengthCodePage), true) == 0)
                {
                    cellType.MaxLengthCodePage = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.MaxLengthUnit), true) == 0)
                {
                    cellType.MaxLengthUnit = prop.Value.ToObject<LengthUnit>();
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.MaxLineCount), true) == 0)
                {
                    cellType.MaxLineCount = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.Multiline), true) == 0)
                {
                    cellType.Multiline = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.PasswordChar), true) == 0)
                {
                    cellType.PasswordChar = prop.Value.ToObject<char>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.PasswordRevelationMode), true) == 0)
                {
                    cellType.PasswordRevelationMode = prop.Value.ToObject<PasswordRevelationMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ReadOnly), true) == 0)
                {
                    cellType.ReadOnly = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.RecommendedValue), true) == 0)
                {
                    cellType.RecommendedValue = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ScrollBarMode), true) == 0)
                {
                    cellType.ScrollBarMode = prop.Value.ToObject<ScrollBarMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ScrollBars), true) == 0)
                {
                    cellType.ScrollBars = prop.Value.ToObject<ScrollBars>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ShortcutKeys), true) == 0)
                {
                    cellType.ShortcutKeys.Clear();
                    ShortcutDictionaryConverter.Deserialize(cellType.ShortcutKeys, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ShowRecommendedValue), true) == 0)
                {
                    cellType.ShowRecommendedValue = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.ShowTouchToolBar), true) == 0)
                {
                    cellType.ShowTouchToolBar = prop.Value.ToObject<TouchToolBarDisplayOptions>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.SideButtons), true) == 0)
                {
                    cellType.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.Static), true) == 0)
                {
                    cellType.Static = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.TouchContextMenuScale), true) == 0)
                {
                    cellType.TouchContextMenuScale = prop.Value.ToObject<float>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.UseSpreadDropDownButtonRender), true) == 0)
                {
                    cellType.UseSpreadDropDownButtonRender = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.UseSystemPasswordChar), true) == 0)
                {
                    cellType.UseSystemPasswordChar = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcTextBoxCellType.WrapMode), true) == 0)
                {
                    cellType.WrapMode = prop.Value.ToObject<WrapMode>();
                    continue;
                }
            }
        }
    }
}
