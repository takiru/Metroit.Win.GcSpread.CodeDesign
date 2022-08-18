using FarPoint.PluginCalendar.WinForms;
using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using FontConverter = Metroit.Win.GcSpread.CodeDesign.Json.Converters.FontConverter;

namespace Metroit.Win.GcSpread.CodeDesign.Json.CellTypeConverters
{
    /// <summary>
    /// DateTimeCellType のコンバーターを提供します。
    /// </summary>
    internal class DateTimeCellTypeConverter : EditBaseCellTypeConverter
    {
        /// <summary>
        /// 新しい DateTimeCellTypeConverter インスタンスを生成します。
        /// </summary>
        /// <param name="cellType">セルタイプ。</param>
        public DateTimeCellTypeConverter(DateTimeCellType cellType) : base(cellType)
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

            var c = CellType as DateTimeCellType;
            if (c == null)
            {
                return;
            }

            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.AMString))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.AMString), c.AMString));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.Calendar))))
            {
                if (c.Calendar is GregorianCalendar)
                {
                    var gregorianCaleander = (GregorianCalendar)c.Calendar;
                    if (gregorianCaleander.CalendarType == GregorianCalendarTypes.Localized)
                    {
                        propsObj.Add(nameof(DateTimeCellType.Calendar), "GregorianLocalized");
                    } else
                    {
                        propsObj.Add(nameof(DateTimeCellType.Calendar), "GregorianUS");
                    }
                }
                if (c.Calendar is JapaneseCalendar)
                {
                    propsObj.Add(nameof(DateTimeCellType.Calendar), "Japanese");
                }
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarDayBackColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarDayBackColor), ColorTranslator.ToHtml(c.CalendarDayBackColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarDayFont))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarDayFont), FontConverter.Serialize(c.CalendarDayFont)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarDayForeColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarDayForeColor), ColorTranslator.ToHtml(c.CalendarDayForeColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarMonthBackColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarMonthBackColor), ColorTranslator.ToHtml(c.CalendarMonthBackColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarMonthFont))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarMonthFont), FontConverter.Serialize(c.CalendarMonthFont)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarMonthForeColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarMonthForeColor), ColorTranslator.ToHtml(c.CalendarMonthForeColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarMonthHeaderDock))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarMonthHeaderDock), c.CalendarMonthHeaderDock));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarMonthHeaderHeight))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarMonthHeaderHeight), c.CalendarMonthHeaderHeight));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarMonthHeaderStyle))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarMonthHeaderStyle), c.CalendarMonthHeaderStyle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarShowSurroundingDays))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarShowSurroundingDays), c.CalendarShowSurroundingDays));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarSingleLineHeader))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarSingleLineHeader), c.CalendarSingleLineHeader));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarSurroundingDaysColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarSurroundingDaysColor), ColorTranslator.ToHtml(c.CalendarSurroundingDaysColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarWeekdayBackColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarWeekdayBackColor), ColorTranslator.ToHtml(c.CalendarWeekdayBackColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarWeekdayFont))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarWeekdayFont), FontConverter.Serialize(c.CalendarWeekdayFont)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarWeekdayForeColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarWeekdayForeColor), ColorTranslator.ToHtml(c.CalendarWeekdayForeColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarWeekdayHeaderDock))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarWeekdayHeaderDock), c.CalendarWeekdayHeaderDock));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarWeekdayHeaderHeight))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarWeekdayHeaderHeight), c.CalendarWeekdayHeaderHeight));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarYearBackColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarYearBackColor), ColorTranslator.ToHtml(c.CalendarYearBackColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarYearFont))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarYearFont), FontConverter.Serialize(c.CalendarYearFont)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarYearForeColor))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarYearForeColor), ColorTranslator.ToHtml(c.CalendarYearForeColor)));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarYearHeaderDock))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarYearHeaderDock), c.CalendarYearHeaderDock));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarYearHeaderHeight))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarYearHeaderHeight), c.CalendarYearHeaderHeight));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.CalendarYearHeaderStyle))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.CalendarYearHeaderStyle), c.CalendarYearHeaderStyle));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.DateDefault))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.DateDefault), c.DateDefault));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.DateSeparator))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.DateSeparator), c.DateSeparator));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.DateTimeFormat))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.DateTimeFormat), c.DateTimeFormat));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.DayNames))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.DayNames), c.DayNames));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.EditorValue))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.EditorValue), c.EditorValue));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.MaximumDate))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.MaximumDate), c.MaximumDate));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.MaximumTime))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.MaximumTime), c.MaximumTime));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.MinimumDate))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.MinimumDate), c.MinimumDate));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.MinimumTime))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.MinimumTime), c.MinimumTime));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.MonthNames))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.MonthNames), c.MonthNames));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.PMString))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.PMString), c.PMString));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.ShortDayNames))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.ShortDayNames), c.ShortDayNames));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.ShortMonthNames))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.ShortMonthNames), c.ShortMonthNames));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.SimpleEdit))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.SimpleEdit), c.SimpleEdit));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.SpinButton))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.SpinButton), c.SpinButton));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.TimeDefault))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.TimeDefault), c.TimeDefault));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.TimeSeparator))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.TimeSeparator), c.TimeSeparator));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.TwoDigitYearMax))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.TwoDigitYearMax), c.TwoDigitYearMax));
            }
            if (includeProps == null || includeProps.Any(x => x.Contains(nameof(DateTimeCellType.UserDefinedFormat))))
            {
                propsObj.Add(new JProperty(nameof(DateTimeCellType.UserDefinedFormat), c.UserDefinedFormat));
            }
        }

        /// <summary>
        /// 現在のプロパティをデシリアライズします。
        /// </summary>
        /// <param name="prop">プロパティのトークンオブジェクト。</param>
        protected override void DeserializeProp(KeyValuePair<string, JToken> prop)
        {
            base.DeserializeProp(prop);

            var c = CellType as DateTimeCellType;
            if (c == null)
            {
                return;
            }

            if (string.Compare(prop.Key, nameof(DateTimeCellType.AMString), true) == 0)
            {
                c.AMString = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.Calendar), true) == 0)
            {
                var propValue = prop.Value.ToString();
                if (string.Compare(propValue, "GregorianLocalized", true) == 0)
                {
                    c.Calendar = new GregorianCalendar(GregorianCalendarTypes.Localized);
                }
                if (string.Compare(propValue, "GregorianUS", true) == 0)
                {
                    c.Calendar = new GregorianCalendar(GregorianCalendarTypes.USEnglish);
                }
                if (string.Compare(propValue, "Japanese", true) == 0)
                {
                    c.Calendar = new JapaneseCalendar();
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarDayBackColor), true) == 0)
            {
                c.CalendarDayBackColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarDayFont), true) == 0)
            {
                var fontObj = FontConverter.Deserialize(prop.Value);
                if (fontObj != null)
                {
                    c.CalendarDayFont = fontObj;
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarDayForeColor), true) == 0)
            {
                c.CalendarDayForeColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarMonthBackColor), true) == 0)
            {
                c.CalendarMonthBackColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarMonthFont), true) == 0)
            {
                var fontObj = FontConverter.Deserialize(prop.Value);
                if (fontObj != null)
                {
                    c.CalendarMonthFont = fontObj;
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarMonthForeColor), true) == 0)
            {
                c.CalendarMonthForeColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarMonthHeaderDock), true) == 0)
            {
                c.CalendarMonthHeaderDock = prop.Value.ToObject<HeaderDock>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarMonthHeaderHeight), true) == 0)
            {
                c.CalendarMonthHeaderHeight = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarMonthHeaderStyle), true) == 0)
            {
                c.CalendarMonthHeaderStyle = prop.Value.ToObject<HeaderStyle>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarShowSurroundingDays), true) == 0)
            {
                c.CalendarShowSurroundingDays = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarSingleLineHeader), true) == 0)
            {
                c.CalendarSingleLineHeader = prop.Value.ToObject<SingleLineHeader>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarSurroundingDaysColor), true) == 0)
            {
                c.CalendarSurroundingDaysColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarWeekdayBackColor), true) == 0)
            {
                c.CalendarWeekdayBackColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarWeekdayFont), true) == 0)
            {
                var fontObj = FontConverter.Deserialize(prop.Value);
                if (fontObj != null)
                {
                    c.CalendarWeekdayFont = fontObj;
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarWeekdayForeColor), true) == 0)
            {
                c.CalendarWeekdayForeColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarWeekdayHeaderDock), true) == 0)
            {
                c.CalendarWeekdayHeaderDock = prop.Value.ToObject<WeekdayDock>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarWeekdayHeaderHeight), true) == 0)
            {
                c.CalendarWeekdayHeaderHeight = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarYearBackColor), true) == 0)
            {
                c.CalendarYearBackColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarYearFont), true) == 0)
            {
                var fontObj = FontConverter.Deserialize(prop.Value);
                if (fontObj != null)
                {
                    c.CalendarYearFont = fontObj;
                }
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarYearForeColor), true) == 0)
            {
                c.CalendarYearForeColor = ColorTranslator.FromHtml(prop.Value.ToObject<string>());
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarYearHeaderDock), true) == 0)
            {
                c.CalendarYearHeaderDock = prop.Value.ToObject<HeaderDock>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarYearHeaderHeight), true) == 0)
            {
                c.CalendarYearHeaderHeight = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.CalendarYearHeaderStyle), true) == 0)
            {
                c.CalendarYearHeaderStyle = prop.Value.ToObject<HeaderStyle>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.DateDefault), true) == 0)
            {
                c.DateDefault = prop.Value.ToObject<DateTime>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.DateSeparator), true) == 0)
            {
                c.DateSeparator = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.DateTimeFormat), true) == 0)
            {
                c.DateTimeFormat = prop.Value.ToObject<DateTimeFormat>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.DayNames), true) == 0)
            {
                c.DayNames = prop.Value.ToObject<string[]>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.EditorValue), true) == 0)
            {
                c.EditorValue = prop.Value.ToObject<DateTimeEditorValue>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.MaximumDate), true) == 0)
            {
                c.MaximumDate = prop.Value.ToObject<DateTime>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.MaximumTime), true) == 0)
            {
                c.MaximumTime = prop.Value.ToObject<TimeSpan>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.MinimumDate), true) == 0)
            {
                c.MinimumDate = prop.Value.ToObject<DateTime>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.MinimumTime), true) == 0)
            {
                c.MinimumTime = prop.Value.ToObject<TimeSpan>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.MonthNames), true) == 0)
            {
                c.MonthNames = prop.Value.ToObject<string[]>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.PMString), true) == 0)
            {
                c.PMString = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.ShortDayNames), true) == 0)
            {
                c.ShortDayNames = prop.Value.ToObject<string[]>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.ShortMonthNames), true) == 0)
            {
                c.ShortMonthNames = prop.Value.ToObject<string[]>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.SimpleEdit), true) == 0)
            {
                c.SimpleEdit = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.SpinButton), true) == 0)
            {
                c.SpinButton = prop.Value.ToObject<bool>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.TimeDefault), true) == 0)
            {
                c.TimeDefault = prop.Value.ToObject<DateTime>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.TimeSeparator), true) == 0)
            {
                c.TimeSeparator = prop.Value.ToObject<string>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.TwoDigitYearMax), true) == 0)
            {
                c.TwoDigitYearMax = prop.Value.ToObject<int>();
                return;
            }
            if (string.Compare(prop.Key, nameof(DateTimeCellType.UserDefinedFormat), true) == 0)
            {
                c.UserDefinedFormat = prop.Value.ToObject<string>();
                return;
            }
        }
    }
}
