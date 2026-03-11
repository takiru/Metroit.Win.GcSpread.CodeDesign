namespace Metroit.Win.GcSpread.CodeDesign.Test
{
    partial class SampleForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            GrapeCity.Win.Spread.InputMan.CellType.GcNumberCellType gcNumberCellType1 = new GrapeCity.Win.Spread.InputMan.CellType.GcNumberCellType();
            GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo dropDownButtonInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo();
            GrapeCity.Win.Spread.InputMan.CellType.GcDateTimeCellType gcDateTimeCellType1 = new GrapeCity.Win.Spread.InputMan.CellType.GcDateTimeCellType();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateYearFieldInfo dateYearFieldInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateYearFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo dateLiteralFieldInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateMonthFieldInfo dateMonthFieldInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateMonthFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo dateLiteralFieldInfo2 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateDayFieldInfo dateDayFieldInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateDayFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo dateLiteralFieldInfo3 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateHourFieldInfo dateHourFieldInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateHourFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo dateLiteralFieldInfo4 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateMinuteFieldInfo dateMinuteFieldInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateMinuteFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo dateLiteralFieldInfo5 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateLiteralFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.Fields.DateSecondFieldInfo dateSecondFieldInfo1 = new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateSecondFieldInfo();
            GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo dropDownButtonInfo2 = new GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo();
            GrapeCity.Win.Spread.InputMan.CellType.GcComboBoxCellType gcComboBoxCellType1 = new GrapeCity.Win.Spread.InputMan.CellType.GcComboBoxCellType();
            GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo dropDownButtonInfo3 = new GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo();
            GrapeCity.Win.Spread.InputMan.CellType.GcNumberCellType gcNumberCellType2 = new GrapeCity.Win.Spread.InputMan.CellType.GcNumberCellType();
            GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo dropDownButtonInfo4 = new GrapeCity.Win.Spread.InputMan.CellType.DropDownButtonInfo();
            GrapeCity.Win.Spread.InputMan.CellType.GcTextBoxCellType gcTextBoxCellType1 = new GrapeCity.Win.Spread.InputMan.CellType.GcTextBoxCellType();
            fpSpread1 = new FarPoint.Win.Spread.FpSpread();
            fpSpread1_Sheet1 = new FarPoint.Win.Spread.SheetView();
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            button4 = new Button();
            button5 = new Button();
            button6 = new Button();
            ((System.ComponentModel.ISupportInitialize)fpSpread1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)fpSpread1_Sheet1).BeginInit();
            SuspendLayout();
            // 
            // fpSpread1
            // 
            fpSpread1.AccessibleDescription = "fpSpread1, Sheet1, Row 0, Column 0";
            fpSpread1.AllowColumnMove = true;
            fpSpread1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            fpSpread1.Location = new Point(14, 51);
            fpSpread1.Margin = new Padding(4, 4, 4, 4);
            fpSpread1.MessageBoxCaption = "SPREADデザイナ";
            fpSpread1.Name = "fpSpread1";
            fpSpread1.Sheets.AddRange(new FarPoint.Win.Spread.SheetView[] { fpSpread1_Sheet1 });
            fpSpread1.Size = new Size(818, 389);
            fpSpread1.TabIndex = 0;
            // 
            // fpSpread1_Sheet1
            // 
            fpSpread1_Sheet1.Reset();
            fpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1;
            fpSpread1_Sheet1.SheetName = "Sheet1";
            fpSpread1_Sheet1.ColumnFooter.DefaultStyle.Locked = false;
            fpSpread1_Sheet1.ColumnFooterSheetCornerStyle.Locked = false;
            fpSpread1_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerDefaultEnhanced";
            fpSpread1_Sheet1.ColumnHeader.DefaultStyle.Locked = false;
            gcNumberCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcNumberCellType1.ClearCollection = true;
            gcNumberCellType1.Fields.IntegerPart.MinDigits = 1;
            gcNumberCellType1.Fields.SignSuffix.NegativePattern = "";
            gcNumberCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] { new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.F2, "SetZero"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.Subtract, "SwitchSign"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.OemMinus, "SwitchSign"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.Control | Keys.Return, "ApplyRecommendedValue") });
            gcNumberCellType1.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] { dropDownButtonInfo1 });
            fpSpread1_Sheet1.Columns.Get(3).CellType = gcNumberCellType1;
            gcDateTimeCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcDateTimeCellType1.ClearCollection = true;
            dateLiteralFieldInfo1.Text = "/";
            dateLiteralFieldInfo2.Text = "/";
            dateLiteralFieldInfo4.Text = ":";
            dateLiteralFieldInfo5.Text = ":";
            gcDateTimeCellType1.Fields.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateFieldInfo[] { dateYearFieldInfo1, dateLiteralFieldInfo1, dateMonthFieldInfo1, dateLiteralFieldInfo2, dateDayFieldInfo1, dateLiteralFieldInfo3, dateHourFieldInfo1, dateLiteralFieldInfo4, dateMinuteFieldInfo1, dateLiteralFieldInfo5, dateSecondFieldInfo1 });
            gcDateTimeCellType1.SerializationDefaultActiveFieldIndex = 0;
            gcDateTimeCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] { new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.F2, "ShortcutClear"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.F5, "SetNow"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.Control | Keys.Return, "ApplyRecommendedValue") });
            gcDateTimeCellType1.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] { dropDownButtonInfo2 });
            fpSpread1_Sheet1.Columns.Get(4).CellType = gcDateTimeCellType1;
            gcComboBoxCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcComboBoxCellType1.ClearCollection = true;
            gcComboBoxCellType1.DropDownMaxHeight = null;
            gcComboBoxCellType1.EllipsisString = "...";
            gcComboBoxCellType1.ListHeaderPane.Height = 22;
            gcComboBoxCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] { new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.F2, "ShortcutClear"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.Control | Keys.Return, "ApplyRecommendedValue") });
            gcComboBoxCellType1.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] { dropDownButtonInfo3 });
            fpSpread1_Sheet1.Columns.Get(5).CellType = gcComboBoxCellType1;
            gcNumberCellType2.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcNumberCellType2.ClearCollection = true;
            gcNumberCellType2.Fields.IntegerPart.MinDigits = 1;
            gcNumberCellType2.Fields.SignSuffix.NegativePattern = "";
            gcNumberCellType2.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] { new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.F2, "SetZero"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.Subtract, "SwitchSign"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.OemMinus, "SwitchSign"), new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.Control | Keys.Return, "ApplyRecommendedValue") });
            gcNumberCellType2.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] { dropDownButtonInfo4 });
            fpSpread1_Sheet1.Columns.Get(6).CellType = gcNumberCellType2;
            gcTextBoxCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcTextBoxCellType1.ClearCollection = true;
            gcTextBoxCellType1.FormatString = "A9N@^ｰﾞﾟ!\"\\#$%&'()=\\^~\\\\|\\@`[{;+:*]},<.>/?";
            gcTextBoxCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] { new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(Keys.F2, "ShortcutClear") });
            fpSpread1_Sheet1.Columns.Get(7).CellType = gcTextBoxCellType1;
            fpSpread1_Sheet1.DefaultStyle.Locked = false;
            fpSpread1_Sheet1.DefaultStyle.Parent = "";
            fpSpread1_Sheet1.FilterBar.DefaultStyle.Locked = false;
            fpSpread1_Sheet1.FilterBar.DefaultStyle.Parent = "";
            fpSpread1_Sheet1.FilterBarHeaderStyle.Locked = false;
            fpSpread1_Sheet1.FilterBarHeaderStyle.Parent = "";
            fpSpread1_Sheet1.RowHeader.Columns.Default.Resizable = false;
            fpSpread1_Sheet1.RowHeader.DefaultStyle.Locked = false;
            fpSpread1_Sheet1.SheetCornerStyle.Locked = false;
            fpSpread1_Sheet1.SheetCornerStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1;
            // 
            // button1
            // 
            button1.Location = new Point(14, 15);
            button1.Margin = new Padding(4, 4, 4, 4);
            button1.Name = "button1";
            button1.Size = new Size(88, 29);
            button1.TabIndex = 1;
            button1.Text = "Bind";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(108, 15);
            button2.Margin = new Padding(4, 4, 4, 4);
            button2.Name = "button2";
            button2.Size = new Size(128, 29);
            button2.TabIndex = 2;
            button2.Text = "SpecialColumJson";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Location = new Point(244, 15);
            button3.Margin = new Padding(4, 4, 4, 4);
            button3.Name = "button3";
            button3.Size = new Size(176, 29);
            button3.TabIndex = 3;
            button3.Text = "SpecialHeaderColumJson";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(427, 15);
            button4.Margin = new Padding(4, 4, 4, 4);
            button4.Name = "button4";
            button4.Size = new Size(94, 29);
            button4.TabIndex = 4;
            button4.Text = "Json";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // button5
            // 
            button5.Location = new Point(528, 15);
            button5.Margin = new Padding(4, 4, 4, 4);
            button5.Name = "button5";
            button5.Size = new Size(134, 29);
            button5.TabIndex = 5;
            button5.Text = "GcComboBoxValue";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // button6
            // 
            button6.Location = new Point(713, 15);
            button6.Name = "button6";
            button6.Size = new Size(75, 23);
            button6.TabIndex = 6;
            button6.Text = "button6";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // SampleForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(847, 455);
            Controls.Add(button6);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(button1);
            Controls.Add(fpSpread1);
            Margin = new Padding(4, 4, 4, 4);
            Name = "SampleForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)fpSpread1).EndInit();
            ((System.ComponentModel.ISupportInitialize)fpSpread1_Sheet1).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private FarPoint.Win.Spread.FpSpread fpSpread1;
        private FarPoint.Win.Spread.SheetView fpSpread1_Sheet1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private Button button6;
    }
}
