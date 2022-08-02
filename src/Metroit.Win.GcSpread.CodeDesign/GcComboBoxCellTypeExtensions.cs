using FarPoint.Win.SuperEdit;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign
{
    /// <summary>
    /// GcComboBoxCellType のJSONシリアライズ、デシリアライズを提供します。
    /// </summary>
    public static class GcComboBoxCellTypeExtensions
    {
        /// <summary>
        /// GcComboBoxCellType をJSONシリアライズします。
        /// DataSource, ImageList, Items, ListColumns, SubEditor, TouchToolBar はシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcComboBoxCellType オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        // NOTE: 下記はシリアライズ化対象外。
        //         DataSource: 様々なオブジェクトが指定されるため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         Items: DataSource 指定もしくは直接指定があるため。
        //         ListColumns: DataSource 指定もしくは直接指定があるため。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static string SerializeJson(this GcComboBoxCellType cellType)
        {
            var result = new JObject();
            result.Add(new JProperty(nameof(cellType.AcceptsArrowKeys), cellType.AcceptsArrowKeys));
            result.Add(new JProperty(nameof(cellType.AcceptsCrLf), cellType.AcceptsCrLf));
            result.Add(new JProperty(nameof(cellType.AcceptsTabChar), cellType.AcceptsTabChar));
            result.Add(new JProperty(nameof(cellType.AllowSpace), cellType.AllowSpace));
            result.Add(new JProperty(nameof(cellType.AlternateText), ComboBoxAlternateTextInfoConverter.Serialize(cellType.AlternateText)));
            result.Add(new JProperty(nameof(cellType.AutoComplete), AutoCompleteInfoConverter.Serialize(cellType.AutoComplete)));
            result.Add(new JProperty(nameof(cellType.AutoCompleteCustomSource), cellType.AutoCompleteCustomSource));
            result.Add(new JProperty(nameof(cellType.AutoCompleteMode), cellType.AutoCompleteMode));
            result.Add(new JProperty(nameof(cellType.AutoCompleteSource), cellType.AutoCompleteSource));
            result.Add(new JProperty(nameof(cellType.AutoConvert), cellType.AutoConvert));
            result.Add(new JProperty(nameof(cellType.AutoFilter), AutoFilterInfoConverter.Serialize(cellType.AutoFilter)));
            result.Add(new JProperty(nameof(cellType.AutoGenerateColumns), cellType.AutoGenerateColumns));
            result.Add(new JProperty(nameof(cellType.AutoSelect), cellType.AutoSelect));
            result.Add(new JProperty(nameof(cellType.BackgroundImage), PictureConverter.Serialize(cellType.BackgroundImage)));
            result.Add(new JProperty(nameof(cellType.DataMember), cellType.DataMember));
            result.Add(new JProperty(nameof(cellType.DropDown), ComboDropDownInfoConverter.Serialize(cellType.DropDown)));
            result.Add(new JProperty(nameof(cellType.DropDownMaxHeight), cellType.DropDownMaxHeight));
            result.Add(new JProperty(nameof(cellType.DropDownStyle), cellType.DropDownStyle));
            result.Add(new JProperty(nameof(cellType.EditMode), cellType.EditMode));
            result.Add(new JProperty(nameof(cellType.EditorValue), cellType.EditorValue));
            result.Add(new JProperty(nameof(cellType.Ellipsis), cellType.Ellipsis));
            result.Add(new JProperty(nameof(cellType.EllipsisString), cellType.EllipsisString));
            result.Add(new JProperty(nameof(cellType.ExcelExportFormat), cellType.ExcelExportFormat));
            result.Add(new JProperty(nameof(cellType.ExitOnLastChar), cellType.ExitOnLastChar));
            result.Add(new JProperty(nameof(cellType.FocusPosition), cellType.FocusPosition));
            result.Add(new JProperty(nameof(cellType.FormatString), cellType.FormatString));
            result.Add(new JProperty(nameof(cellType.ImageAlign), cellType.ImageAlign));
            result.Add(new JProperty(nameof(cellType.ImageWidth), cellType.ImageWidth));
            result.Add(new JProperty(nameof(cellType.ListDefaultColumn), DefaultListColumnInfoConverter.Serialize(cellType.ListDefaultColumn)));
            result.Add(new JProperty(nameof(cellType.ListDescriptionFormat), cellType.ListDescriptionFormat));
            result.Add(new JProperty(nameof(cellType.ListDescriptionSubItemIndex), cellType.ListDescriptionSubItemIndex));
            result.Add(new JProperty(nameof(cellType.ListDisabledItemStyle), ItemStyleInfoConverter.Serialize(cellType.ListDisabledItemStyle)));
            result.Add(new JProperty(nameof(cellType.ListGradientEffect), GradientEffectConverter.Serialize(cellType.ListGradientEffect)));
            result.Add(new JProperty(nameof(cellType.ListGridLines), ListGridLinesInfoConverter.Serialize(cellType.ListGridLines)));
            result.Add(new JProperty(nameof(cellType.ListHeaderPane), ListHeaderPaneInfoConverter.Serialize(cellType.ListHeaderPane)));
            result.Add(new JProperty(nameof(cellType.ListItemTemplates), ItemTemplateCollectionInfoConverter.Serialize(cellType.ListItemTemplates)));
            result.Add(new JProperty(nameof(cellType.ListSelectedItemStyle), ItemStyleInfoConverter.Serialize(cellType.ListSelectedItemStyle)));
            result.Add(new JProperty(nameof(cellType.ListSortColumnIndex), cellType.ListSortColumnIndex));
            result.Add(new JProperty(nameof(cellType.MaxDropDownItems), cellType.MaxDropDownItems));
            result.Add(new JProperty(nameof(cellType.MaxLength), cellType.MaxLength));
            result.Add(new JProperty(nameof(cellType.MaxLengthCodePage), cellType.MaxLengthCodePage));
            result.Add(new JProperty(nameof(cellType.MaxLengthUnit), cellType.MaxLengthUnit));
            result.Add(new JProperty(nameof(cellType.PaintByControl), cellType.PaintByControl));
            result.Add(new JProperty(nameof(cellType.ReadOnly), cellType.ReadOnly));
            result.Add(new JProperty(nameof(cellType.RecommendedValue), cellType.RecommendedValue));
            result.Add(new JProperty(nameof(cellType.ScrollBarMode), cellType.ScrollBarMode));
            result.Add(new JProperty(nameof(cellType.ScrollBars), cellType.ScrollBars));
            result.Add(new JProperty(nameof(cellType.ShortcutKeys), ShortcutDictionaryConverter.Serialize(cellType.ShortcutKeys)));
            result.Add(new JProperty(nameof(cellType.ShowItemTip), cellType.ShowItemTip));
            result.Add(new JProperty(nameof(cellType.ShowListBoxImage), cellType.ShowListBoxImage));
            result.Add(new JProperty(nameof(cellType.ShowOverflowTip), cellType.ShowOverflowTip));
            result.Add(new JProperty(nameof(cellType.ShowRecommendedValue), cellType.ShowRecommendedValue));
            result.Add(new JProperty(nameof(cellType.ShowTouchToolBar), cellType.ShowTouchToolBar));
            result.Add(new JProperty(nameof(cellType.SideButtons), SideButtonCollectionInfoConverter.Serialize(cellType.SideButtons)));
            result.Add(new JProperty(nameof(cellType.Spin), SpinConverter.Serialize(cellType.Spin)));
            result.Add(new JProperty(nameof(cellType.Static), cellType.Static));
            result.Add(new JProperty(nameof(cellType.StatusBar), StatusBarInfoConverter.Serialize(cellType.StatusBar)));
            result.Add(new JProperty(nameof(cellType.TextBoxStyle), cellType.TextBoxStyle));
            result.Add(new JProperty(nameof(cellType.TextFormat), cellType.TextFormat));
            result.Add(new JProperty(nameof(cellType.TextSubItemIndex), cellType.TextSubItemIndex));
            result.Add(new JProperty(nameof(cellType.TouchContextMenuScale), cellType.TouchContextMenuScale));
            result.Add(new JProperty(nameof(cellType.UnSelectedImageIndex), cellType.UnSelectedImageIndex));
            result.Add(new JProperty(nameof(cellType.UseCompatibleDrawing), cellType.UseCompatibleDrawing));
            result.Add(new JProperty(nameof(cellType.UseSpreadDropDownButtonRender), cellType.UseSpreadDropDownButtonRender));
            result.Add(new JProperty(nameof(cellType.ValueSubItemIndex), cellType.ValueSubItemIndex));

            return Newtonsoft.Json.JsonConvert.SerializeObject(result);
        }

        /// <summary>
        /// GcComboBoxCellType をJSONデシリアライズします。
        /// DataSource, ImageList, Items, ListColumns, SubEditor, TouchToolBar はデシリアライズを行いません。
        /// </summary>
        /// <param name="cellType">GcComboBoxCellType オブジェクト。</param>
        /// <param name="cellTypeProps">CellType のプロパティのJSON文字列。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         DataSource: 様々なオブジェクトが指定されるため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         Items: DataSource 指定もしくは直接指定があるため。
        //         ListColumns: DataSource 指定もしくは直接指定があるため。
        //         SubEditor: 様々なオブジェクトが指定される、デザイナで指定不可なため。
        //         TouchToolBar: 大量のプロパティがある、デザイナで指定不可なため。また、再帰情報により StackOverFlow になる。
        public static void DeserializeJson(this GcComboBoxCellType cellType, string cellTypeProps)
        {
            if (string.IsNullOrEmpty(cellTypeProps))
            {
                return;
            }

            var props = JObject.Parse(cellTypeProps);
            foreach (var prop in props)
            {
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AcceptsArrowKeys), true) == 0)
                {
                    cellType.AcceptsArrowKeys = prop.Value.ToObject<AcceptsArrowKeys>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AcceptsCrLf), true) == 0)
                {
                    cellType.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AcceptsTabChar), true) == 0)
                {
                    cellType.AcceptsTabChar = prop.Value.ToObject<TabCharMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AllowSpace), true) == 0)
                {
                    cellType.AllowSpace = prop.Value.ToObject<AllowSpace>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AlternateText), true) == 0)
                {
                    ComboBoxAlternateTextInfoConverter.Deserialize(cellType.AlternateText, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoComplete), true) == 0)
                {
                    AutoCompleteInfoConverter.Deserialize(cellType.AutoComplete, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoCompleteCustomSource), true) == 0)
                {
                    cellType.AutoCompleteCustomSource = prop.Value.ToObject<AutoCompleteStringCollection>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoCompleteMode), true) == 0)
                {
                    cellType.AutoCompleteMode = prop.Value.ToObject<AutoCompleteMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoCompleteSource), true) == 0)
                {
                    cellType.AutoCompleteSource = prop.Value.ToObject<AutoCompleteSource>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoConvert), true) == 0)
                {
                    cellType.AutoConvert = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoFilter), true) == 0)
                {
                    AutoFilterInfoConverter.Deserialize(cellType.AutoFilter, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoGenerateColumns), true) == 0)
                {
                    cellType.AutoGenerateColumns = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoSelect), true) == 0)
                {
                    cellType.AutoSelect = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.BackgroundImage), true) == 0)
                {
                    cellType.BackgroundImage = PictureConverter.Deserialize(prop.Value, cellType.BackgroundImage);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DataMember), true) == 0)
                {
                    cellType.DataMember = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DropDown), true) == 0)
                {
                    ComboDropDownInfoConverter.Deserialize(cellType.DropDown, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DropDownMaxHeight), true) == 0)
                {
                    cellType.DropDownMaxHeight = prop.Value.ToObject<int?>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DropDownStyle), true) == 0)
                {
                    cellType.DropDownStyle = prop.Value.ToObject<ComboBoxStyle>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.EditMode), true) == 0)
                {
                    cellType.EditMode = prop.Value.ToObject<EditMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.EditorValue), true) == 0)
                {
                    cellType.EditorValue = prop.Value.ToObject<GcComboBoxEditorValue>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.Ellipsis), true) == 0)
                {
                    cellType.Ellipsis = prop.Value.ToObject<EllipsisMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.EllipsisString), true) == 0)
                {
                    cellType.EllipsisString = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ExcelExportFormat), true) == 0)
                {
                    cellType.ExcelExportFormat = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ExitOnLastChar), true) == 0)
                {
                    cellType.ExitOnLastChar = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.FocusPosition), true) == 0)
                {
                    cellType.FocusPosition = prop.Value.ToObject<EditorBaseFocusCursorPosition>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.FormatString), true) == 0)
                {
                    cellType.FormatString = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ImageAlign), true) == 0)
                {
                    cellType.ImageAlign = prop.Value.ToObject<System.Windows.Forms.HorizontalAlignment>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ImageWidth), true) == 0)
                {
                    cellType.ImageWidth = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDefaultColumn), true) == 0)
                {
                    cellType.ListDefaultColumn = DefaultListColumnInfoConverter.Deserialize(prop.Value, cellType.ListDefaultColumn);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDescriptionFormat), true) == 0)
                {
                    cellType.ListDescriptionFormat = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDescriptionSubItemIndex), true) == 0)
                {
                    cellType.ListDescriptionSubItemIndex = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDisabledItemStyle), true) == 0)
                {
                    ItemStyleInfoConverter.Deserialize(cellType.ListDisabledItemStyle, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListGradientEffect), true) == 0)
                {
                    cellType.ListGradientEffect = GradientEffectConverter.Deserialize(prop.Value, cellType.ListGradientEffect);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListGridLines), true) == 0)
                {
                    cellType.ListGridLines = ListGridLinesInfoConverter.Deserialize(prop.Value, cellType.ListGridLines);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListHeaderPane), true) == 0)
                {
                    ListHeaderPaneInfoConverter.Deserialize(cellType.ListHeaderPane, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListItemTemplates), true) == 0)
                {
                    cellType.ListItemTemplates = ItemTemplateCollectionInfoConverter.Deserialize((JArray)prop.Value, cellType.ListItemTemplates);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListSelectedItemStyle), true) == 0)
                {
                    ItemStyleInfoConverter.Deserialize(cellType.ListSelectedItemStyle, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListSortColumnIndex), true) == 0)
                {
                    cellType.ListSortColumnIndex = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxDropDownItems), true) == 0)
                {
                    cellType.MaxDropDownItems = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxLength), true) == 0)
                {
                    cellType.MaxLength = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxLengthCodePage), true) == 0)
                {
                    cellType.MaxLengthCodePage = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxLengthUnit), true) == 0)
                {
                    cellType.MaxLengthUnit = prop.Value.ToObject<LengthUnit>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.PaintByControl), true) == 0)
                {
                    cellType.PaintByControl = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ReadOnly), true) == 0)
                {
                    cellType.ReadOnly = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.RecommendedValue), true) == 0)
                {
                    cellType.RecommendedValue = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ScrollBarMode), true) == 0)
                {
                    cellType.ScrollBarMode = prop.Value.ToObject<ScrollBarMode>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ScrollBars), true) == 0)
                {
                    cellType.ScrollBars = prop.Value.ToObject<ScrollBars>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShortcutKeys), true) == 0)
                {
                    cellType.ShortcutKeys.Clear();
                    ShortcutDictionaryConverter.Deserialize(cellType.ShortcutKeys, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowItemTip), true) == 0)
                {
                    cellType.ShowItemTip = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowListBoxImage), true) == 0)
                {
                    cellType.ShowListBoxImage = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowOverflowTip), true) == 0)
                {
                    cellType.ShowOverflowTip = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowRecommendedValue), true) == 0)
                {
                    cellType.ShowRecommendedValue = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowTouchToolBar), true) == 0)
                {
                    cellType.ShowTouchToolBar = prop.Value.ToObject<TouchToolBarDisplayOptions>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.SideButtons), true) == 0)
                {
                    cellType.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.Spin), true) == 0)
                {
                    SpinConverter.Deserialize(cellType.Spin, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.Static), true) == 0)
                {
                    cellType.Static = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.StatusBar), true) == 0)
                {
                    StatusBarInfoConverter.Deserialize(cellType.StatusBar, prop.Value);
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.TextBoxStyle), true) == 0)
                {
                    cellType.TextBoxStyle = prop.Value.ToObject<TextBoxStyle>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.TextFormat), true) == 0)
                {
                    cellType.TextFormat = prop.Value.ToObject<string>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.TextSubItemIndex), true) == 0)
                {
                    cellType.TextSubItemIndex = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.TouchContextMenuScale), true) == 0)
                {
                    cellType.TouchContextMenuScale = prop.Value.ToObject<float>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.UnSelectedImageIndex), true) == 0)
                {
                    cellType.UnSelectedImageIndex = prop.Value.ToObject<int>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.UseCompatibleDrawing), true) == 0)
                {
                    cellType.UseCompatibleDrawing = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.UseSpreadDropDownButtonRender), true) == 0)
                {
                    cellType.UseSpreadDropDownButtonRender = prop.Value.ToObject<bool>();
                    continue;
                }
                if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ValueSubItemIndex), true) == 0)
                {
                    cellType.ValueSubItemIndex = prop.Value.ToObject<int>();
                    continue;
                }
            }
        }
    }
}
