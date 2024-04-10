using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using GrapeCity.Win.Spread.InputMan.CellType;
using Metroit.Win.GcSpread.CodeDesign.Json;
using Metroit.Win.GcSpread.Extensions;
using System.ComponentModel;
using System.Data.Common;
using System.Reflection;

namespace Metroit.Win.GcSpread.CodeDesign.Test
{
    public partial class SampleForm : Form
    {
        public SampleForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // NOTE: ����4�͖����łȂ���΂Ȃ�Ȃ�
            fpSpread1.ActiveSheet.AutoGenerateColumns = false;
            fpSpread1.ActiveSheet.DataAutoCellTypes = false;
            fpSpread1.ActiveSheet.DataAutoSizeColumns = false;
            fpSpread1.ActiveSheet.DataAutoHeadings = false;

            // �R���{�{�b�N�X�p�̃A�C�e��
            var comboItems = new BindingList<ComboItem>()
            {
                new ComboItem() { Display = "item1", Value="value1"},
                new ComboItem() { Display = "item2", Value="value2"},
                new ComboItem() { Display = "item3", Value="value3"}
            };

            // �f�[�^�\�[�X
            var dataSource = new BindingList<ListItem>()
            {
                new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row1", Column4="item1", Column5="aaa", Column6="xxx", Column7=1 },
                new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row2", Column4="item2", Column5="bbb", Column6="yyy", Column7=0 },
                new ListItem() { Column1 = DateTime.Today, Column2=123, Column3="Row3", Column4="item3", Column5="ccc", Column6="zzz", Column7=1 }
            };

            // ���C�A�E�g�I�u�W�F�N�g���烌�C�A�E�g���o�C���h����
            var generator = new SheetViewGenerator(fpSpread1.ActiveSheet);
            generator.GenerateLayout(SampleRex.SampleJson,
                SampleRex.Template,
                null,
                (sheet, root2) =>
                {
                    // �����ŁA���[�U�[�P�ʂ̗�ʒu�A�^�C�g���Ȃǂŏ��������邱�Ƃ��ł���B
                    // Columns[].Options �v���p�e�B�� Tag �ɓ����Ȃǂ������ōs���B
                    // ��ʒu�͉��L�R�[�h�ɂ���čs���A��ړ����ʂ�ԋp����K�v������B
                    // ��ړ����Ȃ��ꍇ�͖߂�l�� null �ł悢�B

                    var result = new List<ColumnMoveResult>();

                    // 1���1�s�ڃw�b�_�[�̃^�C�g����ύX����B
                    sheet.ColumnHeader.Cells[0, 0].Value = "�l�ύX";

                    // Options �� Tag �ɓ����
                    foreach (var column in root2.Columns.Select((Item, Index) => new { Item, Index }))
                    {
                        sheet.Columns[column.Index].Tag = column.Item.Options;
                    }

                    // ��ړ����Ɖ���
                    var columnMoves = new List<(int beforeIndex, int afterIndex, int columnCount)>()
                    {
                        (0, 3, 1),
                        (1, 0, 2),
                        (0, 5, 2)
                    };

                    foreach (var columnMove in columnMoves)
                    {
                        // ��ړ�������
                        // DataField �v���p�e�B��ێ�����K�v�����邽�߁AMoveColumnKeepDataField() �̎��s���K�v�B
                        var cmrs = sheet.MoveColumnKeepDataField(columnMove.beforeIndex, columnMove.afterIndex, columnMove.columnCount);

                        // ���Ɉړ����Ă�����̂���������I���W�i���̃C���f�b�N�X��ێ����ď��������āAresult����͍폜����
                        for (var i = 0; i < cmrs.Count; i++)
                        {
                            var exists = result.Where(x => x.AfterColumnIndex == cmrs[i].BeforeColumnIndex).FirstOrDefault();
                            if (exists != null)
                            {
                                cmrs[i] = new ColumnMoveResult(exists.OriginalColumnIndex, cmrs[i].BeforeColumnIndex, cmrs[i].AfterColumnIndex, cmrs[i].DataField);
                                result.Remove(exists);
                            }
                        }
                        // �ړ���������c������
                        foreach (var cmr in cmrs)
                        {
                            result.Add(cmr);
                        }
                    }

                    return result;
                });

            // ComboBox �̃A�C�e����ݒ�
            // ���񂪊m�肵�Ă��Ȃ����߁A BindJsonLayout() ����ɍs���K�v������
            var comboBoxColumn = fpSpread1.ActiveSheet.Columns.Cast<Column>().Where(x => x.DataField == "Column4").FirstOrDefault();
            var comboBoxCellType = (ComboBoxCellType)comboBoxColumn.CellType;
            comboBoxCellType.Items = new string[] { "item1", "item2", "item3" };
            comboBoxCellType.ItemData = new string[] { "value1", "value2", "value3" };

            // GcComboBox �̃A�C�e���ݒ�
            var cellType = fpSpread1.ActiveSheet.Columns[10].CellType as GcComboBoxCellType;

            //cellType.DropDownStyle = ComboBoxStyle.DropDownList;
            //cellType.EditorValue = GcComboBoxEditorValue.Value; // ���ۂ̒l��Value�Ƃ���

            //// ���s�̋��E�����ז��Ȃ̂ŏ���
            //cellType.ListGridLines.HorizontalLines.Style = LineStyle.None;
            //cellType.ListGridLines.VerticalLines.Style = LineStyle.None;

            //cellType.ListHeaderPane.Visible = false;    // �w�b�_�[�񖼂�\�����Ȃ�
            //cellType.DropDown.AutoWidth = true; // ���\���ɔ����A�����������悤�ɂ���
            //cellType.DropDown.AllowResize = false;  // ���[�U�[�ɂ���ăh���b�v�_�E���̃T�C�Y��ύX���邱�Ƃ������Ȃ�
            //cellType.UseCompatibleDrawing = true;   // �h���b�v�_�E�����J���Ă���Ƃ��ɓ��̓G���A�������D�F�ɂȂ�Ȃ��悤�ɂ���

            //cellType.AutoGenerateColumns = false;
            cellType.DataSource = comboItems;

            //cellType.ListDefaultColumn.Visible = false; // �������X�g�\�����Ȃ���Ԃ��f�t�H���g�ɂ���
            //cellType.ListColumns.Cast<ListColumnInfo>().Where(x => x.DataPropertyName == "Display").First().Visible = true;  // �\���������񂾂��\������
            //cellType.ListColumns.Add(new ListColumnInfo() { DataPropertyName = "Display", DataDisplayType = DataDisplayType.Text, Visible = true });
            //cellType.ListColumns.Add(new ListColumnInfo() { DataPropertyName = "Value", DataDisplayType = DataDisplayType.Text, Visible = false });

            // �����̒l�Ɖ�ʂɕ\�������l��ݒ�
            //cellType.ValueSubItemIndex = cellType.ListColumns.IndexOf(cellType.ListColumns.Cast<ListColumnInfo>().Where(x => x.DataPropertyName == "Value").First());
            //cellType.TextSubItemIndex = cellType.ListColumns.IndexOf(cellType.ListColumns.Cast<ListColumnInfo>().Where(x => x.DataPropertyName == "Display").First());
            //cellType.ValueSubItemIndex = 0;
            //cellType.TextSubItemIndex = 1;

            //cellType.ListItemTemplates.Add(new ItemTemplateInfo() { AutoItemHeight = false, BackColor = Color.Red, Font = new Font(FontFamily.GenericSansSerif, 10), Height = 20, Image = Image.FromFile(@"C:\App\sample.png") });
            cellType.BackgroundImage = new FarPoint.Win.Picture(Image.FromFile(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"sample.png")));

            fpSpread1.ActiveSheet.DataSource = dataSource;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var provider = new ColumnGenerator(fpSpread1.ActiveSheet);
            var defs = provider.CreateColumnsDefinitions(null, new[] { nameof(Column.Width), nameof(Column.Visible), nameof(Column.DataField) }, false, false);
            var json = ColumnGenerator.SerializeJson(defs);
            var obj = ColumnGenerator.DeserializeJson(json);
            MessageBox.Show(json);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var provider = new ColumnHeaderGenerator(fpSpread1.ActiveSheet);
            var defs = provider.CreateCellDefinitions(new[] { nameof(Cell.Value) });
            var json = ColumnHeaderGenerator.SerializeCellJson(defs);
            var obj = ColumnHeaderGenerator.DeserializeCellJson(json);

            MessageBox.Show(json);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var generator = new SheetViewGenerator(fpSpread1.ActiveSheet);
            var root = generator.CreateSheetDefinitions((column) =>
            {
                // �����ŃI�v�V�������̐ݒ�Ȃǂ�����
                // �I�v�V�������Ɖ���
                var options = new List<(string DataField, bool UseFreeWords, bool ExtendWordsEnabled)?>()
                    {
                        ("Column3", false, false),
                        ("Column6", false, false),
                        ("Column1", true, true),
                        ("Column2", true, true),
                        ("Column4", true, true),
                        ("Column5", false, false)
                    };

                var result = new Dictionary<string, object>();

                var option = options.Where(x => x.Value.DataField == column.DataField).FirstOrDefault();
                if (option == null)
                {
                    return result;
                }
                result.Add("UseFreeWords", option.Value.UseFreeWords);
                result.Add("ExtendWordsEnabled", option.Value.ExtendWordsEnabled);

                return result;
            });
            var json = SheetViewGenerator.SerializeJson(root);
            var obj = SheetViewGenerator.DeserializeJson(json);
            MessageBox.Show(json);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var value = fpSpread1.ActiveSheet.Cells[0, 10].Value;
            MessageBox.Show($"{value}");
        }
    }

    public class ListItem
    {
        public DateTime Column1 { get; set; }

        public long Column2 { get; set; }

        public string Column3 { get; set; }

        public string Column4 { get; set; }

        public string Column5 { get; set; }

        public string Column6 { get; set; }

        public int Column7 { get; set; }
    }

    public class ComboItem
    {
        public string Display { get; set; }

        public string Value { get; set; }
    }
}
