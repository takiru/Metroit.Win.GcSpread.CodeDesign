using GrapeCity.Win.Spread.InputMan.CellType;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.Windows.Forms;
using DayOfWeek = GrapeCity.Win.Spread.InputMan.CellType.DayOfWeek;

namespace Metroit.Win.GcSpread.CodeDesign.Json.Converters
{
    /// <summary>
    /// DropDownCalendarInfo オブジェクトのコンバーターを提供します。
    /// </summary>
    internal static class DropDownCalendarInfoConverter
    {
        /// <summary>
        /// シリアライズを行います。
        /// </summary>
        /// <param name="dropDownCalendarInfo">DropDownCalendarInfo オブジェクト。</param>
        /// <returns>シリアライズオブジェクト。</returns>
        public static JObject Serialize(DropDownCalendarInfo dropDownCalendarInfo)
        {
            if (dropDownCalendarInfo == null)
            {
                return null;
            }

            var jObj = new JObject();

            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ActiveHolidayStyles), dropDownCalendarInfo.ActiveHolidayStyles));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.AllowSelection), dropDownCalendarInfo.AllowSelection));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.BackColor), ColorTranslator.ToHtml(dropDownCalendarInfo.BackColor)));

            ImageConverter imgConverter = new ImageConverter();
            byte[] imageBytes = (byte[])imgConverter.ConvertTo(dropDownCalendarInfo.BackgroundImage, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(DropDownCalendarInfo.BackgroundImage), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(DropDownCalendarInfo.BackgroundImage), System.Convert.ToBase64String(imageBytes)));
            }

            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.BackgroundImageLayout), dropDownCalendarInfo.BackgroundImageLayout));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.BorderStyle), dropDownCalendarInfo.BorderStyle));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.CalendarDimensions),
                $"{dropDownCalendarInfo.CalendarDimensions.Width}, {dropDownCalendarInfo.CalendarDimensions.Height}"));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.CalendarMargins),
                $"{dropDownCalendarInfo.CalendarMargins.Left}, {dropDownCalendarInfo.CalendarMargins.Top}, {dropDownCalendarInfo.CalendarMargins.Right}, {dropDownCalendarInfo.CalendarMargins.Bottom}"));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.CalendarType), dropDownCalendarInfo.CalendarType));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ControlStyle), StyleConverter.Serialize(dropDownCalendarInfo.ControlStyle)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.EmptyRows), dropDownCalendarInfo.EmptyRows));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.EnableScrollAnimation), dropDownCalendarInfo.EnableScrollAnimation));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.EnableTouchZoom), dropDownCalendarInfo.EnableTouchZoom));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.FirstDayOfWeek), dropDownCalendarInfo.FirstDayOfWeek));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.FirstMonthInView), dropDownCalendarInfo.FirstMonthInView));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.FlatStyle), dropDownCalendarInfo.FlatStyle));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.FocusRectangle), dropDownCalendarInfo.FocusRectangle));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.Font), FontConverter.Serialize(dropDownCalendarInfo.Font)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.HeaderFormat), dropDownCalendarInfo.HeaderFormat));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.HeaderHeight), dropDownCalendarInfo.HeaderHeight));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.HeaderStyle), StyleConverter.Serialize(dropDownCalendarInfo.HeaderStyle)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.HolidayStyles), HolidayStyleCollectionConverter.Serialize(dropDownCalendarInfo.HolidayStyles)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.InnerMargins),
                $"{dropDownCalendarInfo.InnerMargins.Left}, {dropDownCalendarInfo.InnerMargins.Top}, {dropDownCalendarInfo.InnerMargins.Right}, {dropDownCalendarInfo.InnerMargins.Bottom}"));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.InnerSpace),
                $"{dropDownCalendarInfo.InnerSpace.Width}, {dropDownCalendarInfo.InnerSpace.Height}"));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ItemStyle), StyleConverter.Serialize(dropDownCalendarInfo.ItemStyle)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.LegendStyle), StyleConverter.Serialize(dropDownCalendarInfo.LegendStyle)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.Lines), GridConverter.Serialize(dropDownCalendarInfo.Lines)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.MaxDate), dropDownCalendarInfo.MaxDate));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.MinDate), dropDownCalendarInfo.MinDate));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.NavigateOnWheel), dropDownCalendarInfo.NavigateOnWheel));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.NavigatorOrientation), dropDownCalendarInfo.NavigatorOrientation));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ScrollRate), dropDownCalendarInfo.ScrollRate));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ScrollTipAlign), dropDownCalendarInfo.ScrollTipAlign));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.SelectionStyle), SubStyleConverter.Serialize(dropDownCalendarInfo.SelectionStyle)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowHeader), dropDownCalendarInfo.ShowHeader));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowNavigator), dropDownCalendarInfo.ShowNavigator));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowScrollTip), dropDownCalendarInfo.ShowScrollTip));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowToday), dropDownCalendarInfo.ShowToday));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowTodayMark), dropDownCalendarInfo.ShowTodayMark));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowTrailing), dropDownCalendarInfo.ShowTrailing));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowWeekNumber), dropDownCalendarInfo.ShowWeekNumber));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.ShowZoomButton), dropDownCalendarInfo.ShowZoomButton));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.SingleBorderColor), ColorTranslator.ToHtml(dropDownCalendarInfo.SingleBorderColor)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.TipInterval), dropDownCalendarInfo.TipInterval));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.TitleStyle), StyleConverter.Serialize(dropDownCalendarInfo.TitleStyle)));

            imageBytes = (byte[])imgConverter.ConvertTo(dropDownCalendarInfo.TodayImage, typeof(byte[]));
            if (imageBytes.Length == 0)
            {
                jObj.Add(new JProperty(nameof(DropDownCalendarInfo.TodayImage), null));
            }
            else
            {
                jObj.Add(new JProperty(nameof(DropDownCalendarInfo.TodayImage), System.Convert.ToBase64String(imageBytes)));
            }

            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.TodayMarkColor), ColorTranslator.ToHtml(dropDownCalendarInfo.TodayMarkColor)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.TrailingStyle), SubStyleConverter.Serialize(dropDownCalendarInfo.TrailingStyle)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.UseControlStyle), dropDownCalendarInfo.UseControlStyle));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.UseHeaderFormat), dropDownCalendarInfo.UseHeaderFormat));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.Weekdays), WeekdaysStyleConverter.Serialize(dropDownCalendarInfo.Weekdays)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.WeekNumberStyle), StyleConverter.Serialize(dropDownCalendarInfo.WeekNumberStyle)));
            jObj.Add(new JProperty(nameof(DropDownCalendarInfo.YearMonthFormat), YearMonthFormatConverter.Serialize(dropDownCalendarInfo.YearMonthFormat)));

            return jObj;
        }

        /// <summary>
        /// デシリアライズを行います。
        /// </summary>
        /// <param name="dropDownCalendarInfo">DropDownCalendarInfo オブジェクト。</param>
        /// <param name="prop">デシリアライズオブジェクト。</param>
        public static void Deserialize(DropDownCalendarInfo dropDownCalendarInfo, JToken prop)
        {
            var activeHolidayStyles = prop.SelectToken(nameof(DropDownCalendarInfo.ActiveHolidayStyles));
            if (activeHolidayStyles != null)
            {
                dropDownCalendarInfo.ActiveHolidayStyles = activeHolidayStyles.ToObject<string[]>();
            }

            var allowSelection = prop.SelectToken(nameof(DropDownCalendarInfo.AllowSelection));
            if (allowSelection != null)
            {
                dropDownCalendarInfo.AllowSelection = allowSelection.ToObject<AllowSelection>();
            }

            var backColor = prop.SelectToken(nameof(DropDownCalendarInfo.BackColor));
            if (backColor != null)
            {
                dropDownCalendarInfo.BackColor = ColorTranslator.FromHtml(backColor.ToObject<string>());
            }

            var backgroundImage = prop.SelectToken(nameof(DropDownCalendarInfo.BackgroundImage));
            if (backgroundImage != null)
            {
                if (backgroundImage.HasValues)
                {
                    var imageValue = Convert.FromBase64String(backgroundImage.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    dropDownCalendarInfo.BackgroundImage = imageObj;
                }
                else
                {
                    dropDownCalendarInfo.BackgroundImage = null;
                }
            }

            var backgroundImageLayout = prop.SelectToken(nameof(DropDownCalendarInfo.BackgroundImageLayout));
            if (backgroundImageLayout != null)
            {
                dropDownCalendarInfo.BackgroundImageLayout = backgroundImageLayout.ToObject<ImageLayout>();
            }

            var borderStyle = prop.SelectToken(nameof(DropDownCalendarInfo.BorderStyle));
            if (borderStyle != null)
            {
                dropDownCalendarInfo.BorderStyle = borderStyle.ToObject<BorderStyle>();
            }

            var calendarDimensions = prop.SelectToken(nameof(DropDownCalendarInfo.CalendarDimensions));
            if (calendarDimensions != null)
            {
                dropDownCalendarInfo.CalendarDimensions = calendarDimensions.ToObject<Size>();
            }

            var calendarMargins = prop.SelectToken(nameof(DropDownCalendarInfo.CalendarMargins));
            if (calendarMargins != null)
            {
                dropDownCalendarInfo.CalendarMargins = calendarMargins.ToObject<Padding>();
            }

            var calendarType = prop.SelectToken(nameof(DropDownCalendarInfo.CalendarType));
            if (calendarType != null)
            {
                dropDownCalendarInfo.CalendarType = calendarType.ToObject<CalendarType>();
            }

            var controlStyle = prop.SelectToken(nameof(DropDownCalendarInfo.ControlStyle));
            if (controlStyle != null)
            {
                dropDownCalendarInfo.ControlStyle = StyleConverter.Deserialize(controlStyle, dropDownCalendarInfo.ControlStyle);
            }

            var emptyRows = prop.SelectToken(nameof(DropDownCalendarInfo.EmptyRows));
            if (emptyRows != null)
            {
                dropDownCalendarInfo.EmptyRows = emptyRows.ToObject<EmptyRows>();
            }

            var eenableScrollAnimation = prop.SelectToken(nameof(DropDownCalendarInfo.EnableScrollAnimation));
            if (eenableScrollAnimation != null)
            {
                dropDownCalendarInfo.EnableScrollAnimation = eenableScrollAnimation.ToObject<bool>();
            }

            var enableTouchZoom = prop.SelectToken(nameof(DropDownCalendarInfo.EnableTouchZoom));
            if (enableTouchZoom != null)
            {
                dropDownCalendarInfo.EnableTouchZoom = enableTouchZoom.ToObject<bool>();
            }

            var firstDayOfWeek = prop.SelectToken(nameof(DropDownCalendarInfo.FirstDayOfWeek));
            if (firstDayOfWeek != null)
            {
                dropDownCalendarInfo.FirstDayOfWeek = firstDayOfWeek.ToObject<DayOfWeek>();
            }

            var firstMonthInView = prop.SelectToken(nameof(DropDownCalendarInfo.FirstMonthInView));
            if (firstMonthInView != null)
            {
                dropDownCalendarInfo.FirstMonthInView = firstMonthInView.ToObject<Months>();
            }

            var flatStyle = prop.SelectToken(nameof(DropDownCalendarInfo.FlatStyle));
            if (flatStyle != null)
            {
                dropDownCalendarInfo.FlatStyle = flatStyle.ToObject<FlatStyle>();
            }

            var focusRectangle = prop.SelectToken(nameof(DropDownCalendarInfo.FocusRectangle));
            if (focusRectangle != null)
            {
                dropDownCalendarInfo.FocusRectangle = focusRectangle.ToObject<FocusRectangle>();
            }

            var font = prop.SelectToken(nameof(DropDownCalendarInfo.Font));
            if (font != null)
            {
                var fontObj = FontConverter.Deserialize(font);
                if (fontObj != null)
                {
                    dropDownCalendarInfo.Font = fontObj;
                }
            }

            var headerFormat = prop.SelectToken(nameof(DropDownCalendarInfo.HeaderFormat));
            if (headerFormat != null)
            {
                dropDownCalendarInfo.HeaderFormat = headerFormat.ToObject<string>();
            }

            var headerHeight = prop.SelectToken(nameof(DropDownCalendarInfo.HeaderHeight));
            if (headerHeight != null)
            {
                dropDownCalendarInfo.HeaderHeight = headerHeight.ToObject<int>();
            }

            var headerStyle = prop.SelectToken(nameof(DropDownCalendarInfo.HeaderStyle));
            if (headerStyle != null)
            {
                dropDownCalendarInfo.HeaderStyle = StyleConverter.Deserialize(headerStyle, dropDownCalendarInfo.HeaderStyle);
            }

            var holidayStyles = prop.SelectToken(nameof(DropDownCalendarInfo.HolidayStyles));
            if (holidayStyles != null)
            {
                dropDownCalendarInfo.HolidayStyles = HolidayStyleCollectionConverter.Deserialize((JArray)holidayStyles, dropDownCalendarInfo.HolidayStyles);
            }

            var innerMargins = prop.SelectToken(nameof(DropDownCalendarInfo.InnerMargins));
            if (innerMargins != null)
            {
                dropDownCalendarInfo.InnerMargins = innerMargins.ToObject<Padding>();
            }

            var innerSpace = prop.SelectToken(nameof(DropDownCalendarInfo.InnerSpace));
            if (innerSpace != null)
            {
                dropDownCalendarInfo.InnerSpace = innerSpace.ToObject<Size>();
            }

            var itemStyle = prop.SelectToken(nameof(DropDownCalendarInfo.ItemStyle));
            if (itemStyle != null)
            {
                dropDownCalendarInfo.ItemStyle = StyleConverter.Deserialize(itemStyle, dropDownCalendarInfo.ItemStyle);
            }

            var legendStyle = prop.SelectToken(nameof(DropDownCalendarInfo.LegendStyle));
            if (legendStyle != null)
            {
                dropDownCalendarInfo.LegendStyle = StyleConverter.Deserialize(legendStyle, dropDownCalendarInfo.LegendStyle);
            }

            var lines = prop.SelectToken(nameof(DropDownCalendarInfo.Lines));
            if (lines != null)
            {
                dropDownCalendarInfo.Lines = GridConverter.Deserialize(lines, dropDownCalendarInfo.Lines);
            }

            var maxDate = prop.SelectToken(nameof(DropDownCalendarInfo.MaxDate));
            if (maxDate != null)
            {
                dropDownCalendarInfo.MaxDate = maxDate.ToObject<DateTime?>();
            }

            var minDate = prop.SelectToken(nameof(DropDownCalendarInfo.MinDate));
            if (minDate != null)
            {
                dropDownCalendarInfo.MinDate = minDate.ToObject<DateTime?>();
            }

            var navigateOnWheel = prop.SelectToken(nameof(DropDownCalendarInfo.NavigateOnWheel));
            if (navigateOnWheel != null)
            {
                dropDownCalendarInfo.NavigateOnWheel = navigateOnWheel.ToObject<bool>();
            }

            var nnavigatorOrientation = prop.SelectToken(nameof(DropDownCalendarInfo.NavigatorOrientation));
            if (nnavigatorOrientation != null)
            {
                dropDownCalendarInfo.NavigatorOrientation = nnavigatorOrientation.ToObject<NavigatorOrientation>();
            }

            var scrollRate = prop.SelectToken(nameof(DropDownCalendarInfo.ScrollRate));
            if (scrollRate != null)
            {
                dropDownCalendarInfo.ScrollRate = scrollRate.ToObject<int>();
            }

            var scrollTipAlign = prop.SelectToken(nameof(DropDownCalendarInfo.ScrollTipAlign));
            if (scrollTipAlign != null)
            {
                dropDownCalendarInfo.ScrollTipAlign = scrollTipAlign.ToObject<ScrollTipAlignment>();
            }

            var selectionStyle = prop.SelectToken(nameof(DropDownCalendarInfo.SelectionStyle));
            if (selectionStyle != null)
            {
                dropDownCalendarInfo.SelectionStyle = SubStyleConverter.Deserialize(selectionStyle, dropDownCalendarInfo.SelectionStyle);
            }

            var showHeader = prop.SelectToken(nameof(DropDownCalendarInfo.ShowHeader));
            if (showHeader != null)
            {
                dropDownCalendarInfo.ShowHeader = showHeader.ToObject<bool>();
            }

            var showNavigator = prop.SelectToken(nameof(DropDownCalendarInfo.ShowNavigator));
            if (showNavigator != null)
            {
                dropDownCalendarInfo.ShowNavigator = showNavigator.ToObject<CalendarNavigators>();
            }

            var showScrollTip = prop.SelectToken(nameof(DropDownCalendarInfo.ShowScrollTip));
            if (showScrollTip != null)
            {
                dropDownCalendarInfo.ShowScrollTip = showScrollTip.ToObject<bool>();
            }

            var showToday = prop.SelectToken(nameof(DropDownCalendarInfo.ShowToday));
            if (showToday != null)
            {
                dropDownCalendarInfo.ShowToday = showToday.ToObject<bool>();
            }

            var showTodayMark = prop.SelectToken(nameof(DropDownCalendarInfo.ShowTodayMark));
            if (showTodayMark != null)
            {
                dropDownCalendarInfo.ShowTodayMark = showTodayMark.ToObject<bool>();
            }

            var showTrailing = prop.SelectToken(nameof(DropDownCalendarInfo.ShowTrailing));
            if (showTrailing != null)
            {
                dropDownCalendarInfo.ShowTrailing = showTrailing.ToObject<bool>();
            }

            var showWeekNumber = prop.SelectToken(nameof(DropDownCalendarInfo.ShowWeekNumber));
            if (showWeekNumber != null)
            {
                dropDownCalendarInfo.ShowWeekNumber = showWeekNumber.ToObject<bool>();
            }

            var showZoomButton = prop.SelectToken(nameof(DropDownCalendarInfo.ShowZoomButton));
            if (showZoomButton != null)
            {
                dropDownCalendarInfo.ShowZoomButton = showZoomButton.ToObject<bool>();
            }

            var singleBorderColor = prop.SelectToken(nameof(DropDownCalendarInfo.SingleBorderColor));
            if (singleBorderColor != null)
            {
                dropDownCalendarInfo.SingleBorderColor = ColorTranslator.FromHtml(singleBorderColor.ToObject<string>());
            }

            var tipInterval = prop.SelectToken(nameof(DropDownCalendarInfo.TipInterval));
            if (tipInterval != null)
            {
                dropDownCalendarInfo.TipInterval = tipInterval.ToObject<int>();
            }

            var titleStyle = prop.SelectToken(nameof(DropDownCalendarInfo.TitleStyle));
            if (titleStyle != null)
            {
                dropDownCalendarInfo.TitleStyle = StyleConverter.Deserialize(titleStyle, dropDownCalendarInfo.TitleStyle);
            }

            var todayImage = prop.SelectToken(nameof(DropDownCalendarInfo.TodayImage));
            if (todayImage != null)
            {
                if (todayImage.HasValues)
                {
                    var imageValue = Convert.FromBase64String(todayImage.ToObject<string>());
                    ImageConverter imgConverter = new ImageConverter();
                    var imageObj = (Image)imgConverter.ConvertFrom(imageValue);
                    dropDownCalendarInfo.TodayImage = imageObj;
                }
                else
                {
                    dropDownCalendarInfo.TodayImage = null;
                }
            }

            var todayMarkColor = prop.SelectToken(nameof(DropDownCalendarInfo.TodayMarkColor));
            if (todayMarkColor != null)
            {
                dropDownCalendarInfo.TodayMarkColor = ColorTranslator.FromHtml(todayMarkColor.ToObject<string>());
            }

            var trailingStyle = prop.SelectToken(nameof(DropDownCalendarInfo.TrailingStyle));
            if (trailingStyle != null)
            {
                dropDownCalendarInfo.TrailingStyle = SubStyleConverter.Deserialize(trailingStyle, dropDownCalendarInfo.TrailingStyle);
            }

            var useControlStyle = prop.SelectToken(nameof(DropDownCalendarInfo.UseControlStyle));
            if (useControlStyle != null)
            {
                dropDownCalendarInfo.UseControlStyle = useControlStyle.ToObject<bool>();
            }

            var useHeaderFormat = prop.SelectToken(nameof(DropDownCalendarInfo.UseHeaderFormat));
            if (useHeaderFormat != null)
            {
                dropDownCalendarInfo.UseHeaderFormat = useHeaderFormat.ToObject<bool>();
            }

            var weekdays = prop.SelectToken(nameof(DropDownCalendarInfo.Weekdays));
            if (weekdays != null)
            {
                dropDownCalendarInfo.Weekdays = WeekdaysStyleConverter.Deserialize(weekdays, dropDownCalendarInfo.Weekdays);
            }

            var weekNumberStyle = prop.SelectToken(nameof(DropDownCalendarInfo.WeekNumberStyle));
            if (weekNumberStyle != null)
            {
                dropDownCalendarInfo.WeekNumberStyle = StyleConverter.Deserialize(weekNumberStyle, dropDownCalendarInfo.WeekNumberStyle);
            }

            var yearMonthFormat = prop.SelectToken(nameof(DropDownCalendarInfo.YearMonthFormat));
            if (yearMonthFormat != null)
            {
                dropDownCalendarInfo.YearMonthFormat = YearMonthFormatConverter.Deserialize(yearMonthFormat, dropDownCalendarInfo.YearMonthFormat);
            }
        }
    }
}
