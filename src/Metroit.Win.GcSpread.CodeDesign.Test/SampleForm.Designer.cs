
namespace Metroit.Win.GcSpread.CodeDesign.Test
{
    partial class SampleForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
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
            this.fpSpread1 = new FarPoint.Win.Spread.FpSpread();
            this.fpSpread1_Sheet1 = new FarPoint.Win.Spread.SheetView();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.fpSpread1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fpSpread1_Sheet1)).BeginInit();
            this.SuspendLayout();
            // 
            // fpSpread1
            // 
            this.fpSpread1.AccessibleDescription = "fpSpread1, Sheet1, Row 0, Column 0";
            this.fpSpread1.AllowColumnMove = true;
            this.fpSpread1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fpSpread1.Location = new System.Drawing.Point(12, 41);
            this.fpSpread1.Name = "fpSpread1";
            this.fpSpread1.Sheets.AddRange(new FarPoint.Win.Spread.SheetView[] {
            this.fpSpread1_Sheet1});
            this.fpSpread1.Size = new System.Drawing.Size(702, 312);
            this.fpSpread1.TabIndex = 0;
            // 
            // fpSpread1_Sheet1
            // 
            this.fpSpread1_Sheet1.Reset();
            this.fpSpread1_Sheet1.SheetName = "Sheet1";
            // Formulas and custom names must be loaded with R1C1 reference style
            this.fpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1;
            this.fpSpread1_Sheet1.ColumnFooterSheetCornerStyle.BackColor = System.Drawing.Color.Empty;
            this.fpSpread1_Sheet1.ColumnFooterSheetCornerStyle.ForeColor = System.Drawing.Color.Empty;
            this.fpSpread1_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerDefaultEnhanced";
            gcNumberCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, System.Drawing.Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcNumberCellType1.ClearCollection = true;
            gcNumberCellType1.Fields.IntegerPart.MinDigits = 1;
            gcNumberCellType1.Fields.SignSuffix.NegativePattern = "";
            gcNumberCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] {
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.F2, "SetZero"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.Subtract, "SwitchSign"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.OemMinus, "SwitchSign"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Return))), "ApplyRecommendedValue")});
            gcNumberCellType1.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] {
            dropDownButtonInfo1});
            this.fpSpread1_Sheet1.Columns.Get(3).CellType = gcNumberCellType1;
            gcDateTimeCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, System.Drawing.Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcDateTimeCellType1.ClearCollection = true;
            dateLiteralFieldInfo1.Text = "/";
            dateLiteralFieldInfo2.Text = "/";
            dateLiteralFieldInfo4.Text = ":";
            dateLiteralFieldInfo5.Text = ":";
            gcDateTimeCellType1.Fields.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.Fields.DateFieldInfo[] {
            dateYearFieldInfo1,
            dateLiteralFieldInfo1,
            dateMonthFieldInfo1,
            dateLiteralFieldInfo2,
            dateDayFieldInfo1,
            dateLiteralFieldInfo3,
            dateHourFieldInfo1,
            dateLiteralFieldInfo4,
            dateMinuteFieldInfo1,
            dateLiteralFieldInfo5,
            dateSecondFieldInfo1});
            gcDateTimeCellType1.SerializationDefaultActiveFieldIndex = 0;
            gcDateTimeCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] {
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.F2, "ShortcutClear"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.F5, "SetNow"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Return))), "ApplyRecommendedValue")});
            gcDateTimeCellType1.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] {
            dropDownButtonInfo2});
            this.fpSpread1_Sheet1.Columns.Get(4).CellType = gcDateTimeCellType1;
            gcComboBoxCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, System.Drawing.Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcComboBoxCellType1.ClearCollection = true;
            gcComboBoxCellType1.DropDownMaxHeight = null;
            gcComboBoxCellType1.EllipsisString = "...";
            gcComboBoxCellType1.ListHeaderPane.Height = 19;
            gcComboBoxCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] {
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.F2, "ShortcutClear"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Return))), "ApplyRecommendedValue")});
            gcComboBoxCellType1.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] {
            dropDownButtonInfo3});
            this.fpSpread1_Sheet1.Columns.Get(5).CellType = gcComboBoxCellType1;
            gcNumberCellType2.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, System.Drawing.Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcNumberCellType2.ClearCollection = true;
            gcNumberCellType2.Fields.IntegerPart.MinDigits = 1;
            gcNumberCellType2.Fields.SignSuffix.NegativePattern = "";
            gcNumberCellType2.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] {
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.F2, "SetZero"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.Subtract, "SwitchSign"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.OemMinus, "SwitchSign"),
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Return))), "ApplyRecommendedValue")});
            gcNumberCellType2.SideButtons.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.SideButtonBaseInfo[] {
            dropDownButtonInfo4});
            this.fpSpread1_Sheet1.Columns.Get(6).CellType = gcNumberCellType2;
            gcTextBoxCellType1.BackgroundImage = new FarPoint.Win.Picture(null, FarPoint.Win.RenderStyle.Normal, System.Drawing.Color.Empty, 0, FarPoint.Win.HorizontalAlignment.Left, FarPoint.Win.VerticalAlignment.Top);
            gcTextBoxCellType1.ClearCollection = true;
            gcTextBoxCellType1.FormatString = "A9N@^ｰﾞﾟ!\"\\#$%&\'()=\\^~\\\\|\\@`[{;+:*]},<.>/?";
            gcTextBoxCellType1.ShortcutKeys.AddRange(new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry[] {
            new GrapeCity.Win.Spread.InputMan.CellType.ShortcutDictionaryEntry(System.Windows.Forms.Keys.F2, "ShortcutClear")});
            this.fpSpread1_Sheet1.Columns.Get(7).CellType = gcTextBoxCellType1;
            this.fpSpread1_Sheet1.RowHeader.Columns.Default.Resizable = false;
            this.fpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Bind";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(93, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(110, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "SpecialColumJson";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(209, 12);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(151, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "SpecialHeaderColumJson";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(366, 12);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(81, 23);
            this.button4.TabIndex = 4;
            this.button4.Text = "Json";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(453, 12);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(115, 23);
            this.button5.TabIndex = 5;
            this.button5.Text = "GcComboBoxValue";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // SampleForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(726, 364);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.fpSpread1);
            this.Name = "SampleForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.fpSpread1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fpSpread1_Sheet1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private FarPoint.Win.Spread.FpSpread fpSpread1;
        private FarPoint.Win.Spread.SheetView fpSpread1_Sheet1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
    }
}

