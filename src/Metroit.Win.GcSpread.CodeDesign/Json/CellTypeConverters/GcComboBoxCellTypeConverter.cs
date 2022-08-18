using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// GcComboBoxCellType のコンバーターを提供します。
    /// </summary>
    internal class GcComboBoxCellTypeConverter : InputManCellTypeBaseConverter
    {
        /// <summary>
        /// 新しい GcComboBoxCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public GcComboBoxCellTypeConverter(GcComboBoxCellType cellType) : base(cellType)
        {

        }

        /// <summary>
        /// セルタイプのプロパティをシリアライズします。
        /// DataSource, ImageList, Items, ListColumns はシリアライズを行いません。
        /// </summary>
        /// <param name="propsObj">プロパティオブジェクト。</param>
        /// <param name="includeProps">シリアライズに含めるプロパティ名。nullの場合はすべてのプロパティ、指定した場合は指定したプロパティのみがシリアライズされます。</param>
        // NOTE: 下記はシリアライズ化対象外。
        //         DataSource: 様々なオブジェクトが指定されるため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         Items: DataSource 指定もしくは直接指定があるため。
        //         ListColumns: DataSource 指定もしくは直接指定があるため。
        protected override void SerializeProp(JObject propsObj, string[] includeProps)
        {
            base.SerializeProp(propsObj, includeProps);

            var c = CellType as GcComboBoxCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AcceptsCrLf))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AcceptsCrLf), c.AcceptsCrLf));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AcceptsTabChar))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AcceptsTabChar), c.AcceptsTabChar));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AllowSpace))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AllowSpace), c.AllowSpace));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AlternateText))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AlternateText), ComboBoxAlternateTextInfoConverter.Serialize(c.AlternateText)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoComplete))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoComplete), AutoCompleteInfoConverter.Serialize(c.AutoComplete)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoCompleteCustomSource))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoCompleteCustomSource), c.AutoCompleteCustomSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoCompleteMode))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoCompleteMode), c.AutoCompleteMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoCompleteSource))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoCompleteSource), c.AutoCompleteSource));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoConvert))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoConvert), c.AutoConvert));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoFilter))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoFilter), AutoFilterInfoConverter.Serialize(c.AutoFilter)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoGenerateColumns))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoGenerateColumns), c.AutoGenerateColumns));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.AutoSelect))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.AutoSelect), c.AutoSelect));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.DataMember))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.DataMember), c.DataMember));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.DropDown))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.DropDown), ComboDropDownInfoConverter.Serialize(c.DropDown)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.DropDownMaxHeight))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.DropDownMaxHeight), c.DropDownMaxHeight));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.DropDownStyle))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.DropDownStyle), c.DropDownStyle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.EditorValue))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.EditorValue), c.EditorValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.Ellipsis))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.Ellipsis), c.Ellipsis));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.EllipsisString))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.EllipsisString), c.EllipsisString));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.FocusPosition))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.FocusPosition), c.FocusPosition));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.FormatString))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.FormatString), c.FormatString));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ImageAlign))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ImageAlign), c.ImageAlign));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ImageWidth))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ImageWidth), c.ImageWidth));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListDefaultColumn))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListDefaultColumn), DefaultListColumnInfoConverter.Serialize(c.ListDefaultColumn)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListDescriptionFormat))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListDescriptionFormat), c.ListDescriptionFormat));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListDescriptionSubItemIndex))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListDescriptionSubItemIndex), c.ListDescriptionSubItemIndex));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListDisabledItemStyle))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListDisabledItemStyle), ItemStyleInfoConverter.Serialize(c.ListDisabledItemStyle)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListGradientEffect))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListGradientEffect), GradientEffectConverter.Serialize(c.ListGradientEffect)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListGridLines))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListGridLines), ListGridLinesInfoConverter.Serialize(c.ListGridLines)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListHeaderPane))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListHeaderPane), ListHeaderPaneInfoConverter.Serialize(c.ListHeaderPane)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListItemTemplates))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListItemTemplates), ItemTemplateCollectionInfoConverter.Serialize(c.ListItemTemplates)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListSelectedItemStyle))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListSelectedItemStyle), ItemStyleInfoConverter.Serialize(c.ListSelectedItemStyle)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ListSortColumnIndex))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ListSortColumnIndex), c.ListSortColumnIndex));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.MaxDropDownItems))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.MaxDropDownItems), c.MaxDropDownItems));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.MaxLength))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.MaxLength), c.MaxLength));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.MaxLengthCodePage))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.MaxLengthCodePage), c.MaxLengthCodePage));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.MaxLengthUnit))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.MaxLengthUnit), c.MaxLengthUnit));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.PaintByControl))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.PaintByControl), c.PaintByControl));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.RecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.RecommendedValue), c.RecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ScrollBarMode))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ScrollBarMode), c.ScrollBarMode));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ScrollBars))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ScrollBars), c.ScrollBars));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ShowItemTip))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ShowItemTip), c.ShowItemTip));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ShowListBoxImage))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ShowListBoxImage), c.ShowListBoxImage));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ShowOverflowTip))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ShowOverflowTip), c.ShowOverflowTip));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ShowRecommendedValue))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ShowRecommendedValue), c.ShowRecommendedValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.SideButtons))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.SideButtons), SideButtonCollectionInfoConverter.Serialize(c.SideButtons)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.Spin))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.Spin), SpinConverter.Serialize(c.Spin)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.StatusBar))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.StatusBar), StatusBarInfoConverter.Serialize(c.StatusBar)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.TextBoxStyle))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.TextBoxStyle), c.TextBoxStyle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.TextFormat))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.TextFormat), c.TextFormat));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.TextSubItemIndex))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.TextSubItemIndex), c.TextSubItemIndex));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.UnSelectedImageIndex))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.UnSelectedImageIndex), c.UnSelectedImageIndex));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.UseCompatibleDrawing))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.UseCompatibleDrawing), c.UseCompatibleDrawing));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(GcComboBoxCellType.ValueSubItemIndex))))
            {
                propsObj.Add(new JProperty(nameof(GcComboBoxCellType.ValueSubItemIndex), c.ValueSubItemIndex));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// DataSource, ImageList, Items, ListColumns はデシリアライズを行いません。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        // NOTE: 下記はデシリアライズ化対象外。
        //         DataSource: 様々なオブジェクトが指定されるため。
        //         ImageList: 複数のオブジェクトが対象になる、デザイナで指定不可なため。
        //         Items: DataSource 指定もしくは直接指定があるため。
        //         ListColumns: DataSource 指定もしくは直接指定があるため。
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as GcComboBoxCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AcceptsCrLf), true) == 0)
            {
                c.AcceptsCrLf = prop.Value.ToObject<CrLfMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AcceptsTabChar), true) == 0)
            {
                c.AcceptsTabChar = prop.Value.ToObject<TabCharMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AllowSpace), true) == 0)
            {
                c.AllowSpace = prop.Value.ToObject<AllowSpace>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AlternateText), true) == 0)
            {
                ComboBoxAlternateTextInfoConverter.Deserialize(c.AlternateText, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoComplete), true) == 0)
            {
                AutoCompleteInfoConverter.Deserialize(c.AutoComplete, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoCompleteCustomSource), true) == 0)
            {
                c.AutoCompleteCustomSource = prop.Value.ToObject<AutoCompleteStringCollection>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoCompleteMode), true) == 0)
            {
                c.AutoCompleteMode = prop.Value.ToObject<AutoCompleteMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoCompleteSource), true) == 0)
            {
                c.AutoCompleteSource = prop.Value.ToObject<AutoCompleteSource>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoConvert), true) == 0)
            {
                c.AutoConvert = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoFilter), true) == 0)
            {
                AutoFilterInfoConverter.Deserialize(c.AutoFilter, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoGenerateColumns), true) == 0)
            {
                c.AutoGenerateColumns = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.AutoSelect), true) == 0)
            {
                c.AutoSelect = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DataMember), true) == 0)
            {
                c.DataMember = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DropDown), true) == 0)
            {
                ComboDropDownInfoConverter.Deserialize(c.DropDown, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DropDownMaxHeight), true) == 0)
            {
                c.DropDownMaxHeight = prop.Value.ToObject<int?>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.DropDownStyle), true) == 0)
            {
                c.DropDownStyle = prop.Value.ToObject<ComboBoxStyle>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.EditorValue), true) == 0)
            {
                c.EditorValue = prop.Value.ToObject<GcComboBoxEditorValue>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.Ellipsis), true) == 0)
            {
                c.Ellipsis = prop.Value.ToObject<EllipsisMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.EllipsisString), true) == 0)
            {
                c.EllipsisString = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.FocusPosition), true) == 0)
            {
                c.FocusPosition = prop.Value.ToObject<EditorBaseFocusCursorPosition>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.FormatString), true) == 0)
            {
                c.FormatString = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ImageAlign), true) == 0)
            {
                c.ImageAlign = prop.Value.ToObject<System.Windows.Forms.HorizontalAlignment>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ImageWidth), true) == 0)
            {
                c.ImageWidth = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDefaultColumn), true) == 0)
            {
                c.ListDefaultColumn = DefaultListColumnInfoConverter.Deserialize(prop.Value, c.ListDefaultColumn);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDescriptionFormat), true) == 0)
            {
                c.ListDescriptionFormat = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDescriptionSubItemIndex), true) == 0)
            {
                c.ListDescriptionSubItemIndex = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListDisabledItemStyle), true) == 0)
            {
                ItemStyleInfoConverter.Deserialize(c.ListDisabledItemStyle, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListGradientEffect), true) == 0)
            {
                c.ListGradientEffect = GradientEffectConverter.Deserialize(prop.Value, c.ListGradientEffect);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListGridLines), true) == 0)
            {
                c.ListGridLines = ListGridLinesInfoConverter.Deserialize(prop.Value, c.ListGridLines);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListHeaderPane), true) == 0)
            {
                ListHeaderPaneInfoConverter.Deserialize(c.ListHeaderPane, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListItemTemplates), true) == 0)
            {
                c.ListItemTemplates = ItemTemplateCollectionInfoConverter.Deserialize((JArray)prop.Value, c.ListItemTemplates);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListSelectedItemStyle), true) == 0)
            {
                ItemStyleInfoConverter.Deserialize(c.ListSelectedItemStyle, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ListSortColumnIndex), true) == 0)
            {
                c.ListSortColumnIndex = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxDropDownItems), true) == 0)
            {
                c.MaxDropDownItems = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxLength), true) == 0)
            {
                c.MaxLength = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxLengthCodePage), true) == 0)
            {
                c.MaxLengthCodePage = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.MaxLengthUnit), true) == 0)
            {
                c.MaxLengthUnit = prop.Value.ToObject<LengthUnit>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.PaintByControl), true) == 0)
            {
                c.PaintByControl = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.RecommendedValue), true) == 0)
            {
                c.RecommendedValue = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ScrollBarMode), true) == 0)
            {
                c.ScrollBarMode = prop.Value.ToObject<ScrollBarMode>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ScrollBars), true) == 0)
            {
                c.ScrollBars = prop.Value.ToObject<ScrollBars>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowItemTip), true) == 0)
            {
                c.ShowItemTip = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowListBoxImage), true) == 0)
            {
                c.ShowListBoxImage = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowOverflowTip), true) == 0)
            {
                c.ShowOverflowTip = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ShowRecommendedValue), true) == 0)
            {
                c.ShowRecommendedValue = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.SideButtons), true) == 0)
            {
                c.SideButtons = SideButtonCollectionInfoConverter.Deserialize((JArray)prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.Spin), true) == 0)
            {
                SpinConverter.Deserialize(c.Spin, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.StatusBar), true) == 0)
            {
                StatusBarInfoConverter.Deserialize(c.StatusBar, prop.Value);
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.TextBoxStyle), true) == 0)
            {
                c.TextBoxStyle = prop.Value.ToObject<TextBoxStyle>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.TextFormat), true) == 0)
            {
                c.TextFormat = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.TextSubItemIndex), true) == 0)
            {
                c.TextSubItemIndex = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.UnSelectedImageIndex), true) == 0)
            {
                c.UnSelectedImageIndex = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.UseCompatibleDrawing), true) == 0)
            {
                c.UseCompatibleDrawing = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(GcComboBoxCellType.ValueSubItemIndex), true) == 0)
            {
                c.ValueSubItemIndex = prop.Value.ToObject<int>();
                return;
            }
        }
    }
}
