using System;
using System.Drawing;
using System.Drawing.Text;
using System.Windows.Forms;

namespace PowerPointUsefulTools
{
    internal class TableSettingsForm : Form
    {
        private readonly StyleControls _header = new StyleControls();
        private readonly StyleControls _body = new StyleControls();

        public DefaultTableSettings Result { get; private set; }

        public TableSettingsForm(DefaultTableSettings settings)
        {
            BuildUI();
            LoadSettings(settings);
        }

        // ---------------------------------------------------------------
        // UI construction
        // ---------------------------------------------------------------

        private void BuildUI()
        {
            Text = "デフォルトテーブルスタイル設定";
            Size = new Size(780, 640);
            MinimumSize = new Size(780, 640);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(8)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44f));

            root.Controls.Add(BuildStyleGroup("ヘッダースタイル", _header), 0, 0);
            root.Controls.Add(BuildStyleGroup("ボディスタイル", _body), 1, 0);

            var btnFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 6, 4, 0)
            };

            var cancelBtn = new Button { Text = "キャンセル", Width = 90, DialogResult = DialogResult.Cancel };
            var okBtn = new Button { Text = "OK", Width = 90 };
            okBtn.Click += (s, e) =>
            {
                Result = BuildSettings();
                DialogResult = DialogResult.OK;
                Close();
            };
            btnFlow.Controls.Add(cancelBtn);
            btnFlow.Controls.Add(okBtn);

            root.SetColumnSpan(btnFlow, 2);
            root.Controls.Add(btnFlow, 0, 1);

            Controls.Add(root);
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private GroupBox BuildStyleGroup(string title, StyleControls ctrl)
        {
            var grp = new GroupBox
            {
                Text = title,
                Dock = DockStyle.Fill,
                Padding = new Padding(6, 4, 6, 6)
            };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 0
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115f));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            // 背景色
            ctrl.FillColorBtn = CreateColorButton(Color.White);
            ctrl.FillColorBtn.Click += (s, e) => PickColor(ctrl.FillColorBtn, v => ctrl.FillColorRGB = v);
            AddRow(grid, "背景色", ctrl.FillColorBtn);

            // 透明度
            ctrl.TransparencyNud = new NumericUpDown
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 100,
                DecimalPlaces = 0,
                Increment = 1,
                Value = 0
            };
            AddRow(grid, "透明度 (%)", ctrl.TransparencyNud);

            // 文字色
            ctrl.FontColorBtn = CreateColorButton(Color.Black);
            ctrl.FontColorBtn.Click += (s, e) => PickColor(ctrl.FontColorBtn, v => ctrl.FontColorRGB = v);
            AddRow(grid, "文字色", ctrl.FontColorBtn);

            // フォント名
            ctrl.FontNameCmb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDown };
            PopulateFontNames(ctrl.FontNameCmb);
            AddRow(grid, "フォント名", ctrl.FontNameCmb);

            // フォントサイズ
            ctrl.FontSizeNud = new NumericUpDown
            {
                Dock = DockStyle.Fill,
                Minimum = 1,
                Maximum = 200,
                DecimalPlaces = 1,
                Increment = 0.5m,
                Value = 11
            };
            AddRow(grid, "サイズ (pt)", ctrl.FontSizeNud);

            // スタイル
            ctrl.BoldChk = new CheckBox { Text = "太字", AutoSize = true };
            ctrl.ItalicChk = new CheckBox { Text = "斜体", AutoSize = true };
            var styleFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            styleFlow.Controls.Add(ctrl.BoldChk);
            styleFlow.Controls.Add(ctrl.ItalicChk);
            AddRow(grid, "スタイル", styleFlow);

            // 余白
            ctrl.MarginTopNud = CreateMarginNud();
            ctrl.MarginBottomNud = CreateMarginNud();
            ctrl.MarginLeftNud = CreateMarginNud();
            ctrl.MarginRightNud = CreateMarginNud();
            AddRow(grid, "余白 (pt)", BuildMarginPanel(ctrl), ContentAlignment.TopRight);

            // 罫線: ヘッダー行 22pt + データ行 4×28pt = 134pt で固定
            {
                var borderRow = grid.RowStyles.Count;
                grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 134f));
                grid.RowCount = borderRow + 1;
                grid.Controls.Add(new Label
                {
                    Text = "罫線",
                    TextAlign = ContentAlignment.TopRight,
                    Dock = DockStyle.Fill,
                    Padding = new Padding(0, 4, 6, 4)
                }, 0, borderRow);
                grid.Controls.Add(BuildBorderPanel(ctrl), 1, borderRow);
            }

            grp.Controls.Add(grid);
            return grp;
        }

        private static TableLayoutPanel BuildMarginPanel(StyleControls ctrl)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 2,
                AutoSize = true
            };
            for (int i = 0; i < 4; i++)
                panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            panel.Controls.Add(CenterLabel("上"), 0, 0);
            panel.Controls.Add(CenterLabel("下"), 1, 0);
            panel.Controls.Add(CenterLabel("左"), 2, 0);
            panel.Controls.Add(CenterLabel("右"), 3, 0);
            panel.Controls.Add(ctrl.MarginTopNud, 0, 1);
            panel.Controls.Add(ctrl.MarginBottomNud, 1, 1);
            panel.Controls.Add(ctrl.MarginLeftNud, 2, 1);
            panel.Controls.Add(ctrl.MarginRightNud, 3, 1);
            return panel;
        }

        private static TableLayoutPanel BuildBorderPanel(StyleControls ctrl)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                RowCount = 5
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 24f));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 46f));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40f));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60f));

            // RowCount=5 のセッターが追加したデフォルト行をクリアし、正しいスタイルを一括設定
            panel.RowStyles.Clear();
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 22f));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));

            panel.Controls.Add(new Label(), 0, 0);
            panel.Controls.Add(CenterLabel("表示"), 1, 0);
            panel.Controls.Add(CenterLabel("色"), 2, 0);
            panel.Controls.Add(CenterLabel("太さ(pt)"), 3, 0);
            panel.Controls.Add(CenterLabel("線種"), 4, 0);

            ctrl.BorderTop = AddBorderRow(panel, "上", 1);
            ctrl.BorderBottom = AddBorderRow(panel, "下", 2);
            ctrl.BorderLeft = AddBorderRow(panel, "左", 3);
            ctrl.BorderRight = AddBorderRow(panel, "右", 4);
            return panel;
        }

        private static BorderRowControls AddBorderRow(TableLayoutPanel panel, string name, int row)
        {
            var ctrl = new BorderRowControls();

            panel.Controls.Add(CenterLabel(name), 0, row);

            ctrl.VisibleChk = new CheckBox { Dock = DockStyle.Fill, CheckAlign = ContentAlignment.MiddleCenter };
            panel.Controls.Add(ctrl.VisibleChk, 1, row);

            ctrl.ColorBtn = new Button
            {
                BackColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Text = string.Empty,
                Dock = DockStyle.Fill,
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            ctrl.ColorBtn.FlatAppearance.BorderColor = Color.DimGray;
            var captured = ctrl;
            ctrl.ColorBtn.Click += (s, e) => PickColor(captured.ColorBtn, v => captured.ColorRGB = v);
            panel.Controls.Add(ctrl.ColorBtn, 2, row);

            ctrl.WeightNud = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 10,
                DecimalPlaces = 2,
                Increment = 0.25m,
                Value = 0.75m,
                Dock = DockStyle.Fill
            };
            panel.Controls.Add(ctrl.WeightNud, 3, row);

            ctrl.DashCmb = CreateDashCombo();
            panel.Controls.Add(ctrl.DashCmb, 4, row);
            return ctrl;
        }

        private static ComboBox CreateDashCombo()
        {
            var cmb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            cmb.Items.Add("実線");
            cmb.Items.Add("点線");
            cmb.Items.Add("破線");
            cmb.Items.Add("一点鎖線");
            cmb.Items.Add("二点鎖線");
            cmb.SelectedIndex = 0;
            return cmb;
        }

        private static Label CenterLabel(string text)
        {
            return new Label
            {
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };
        }

        private static void AddRow(TableLayoutPanel grid, string labelText, Control control,
            ContentAlignment labelAlign = ContentAlignment.MiddleRight)
        {
            var row = grid.RowStyles.Count;
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            grid.RowCount = row + 1;
            grid.Controls.Add(new Label
            {
                Text = labelText,
                TextAlign = labelAlign,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 4, 6, 4)
            }, 0, row);
            grid.Controls.Add(control, 1, row);
        }

        // ---------------------------------------------------------------
        // Color picker helpers
        // ---------------------------------------------------------------

        private static Button CreateColorButton(Color initial)
        {
            var btn = new Button
            {
                BackColor = initial,
                FlatStyle = FlatStyle.Flat,
                Text = string.Empty,
                Height = 28,
                Width = 70,
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            btn.FlatAppearance.BorderColor = Color.DimGray;
            return btn;
        }

        private static void PickColor(Button colorBtn, Action<int> onSelected)
        {
            using (var dlg = new ColorDialog
            {
                Color = colorBtn.BackColor,
                FullOpen = true,
                AnyColor = true
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                colorBtn.BackColor = dlg.Color;
                colorBtn.FlatAppearance.MouseOverBackColor = dlg.Color;
                onSelected(ColorToOfficeBgr(dlg.Color));
            }
        }

        // ---------------------------------------------------------------
        // Control helpers
        // ---------------------------------------------------------------

        private static NumericUpDown CreateMarginNud()
        {
            return new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                DecimalPlaces = 1,
                Increment = 0.5m,
                Value = 3.6m,
                Dock = DockStyle.Fill
            };
        }

        private static void PopulateFontNames(ComboBox cmb)
        {
            using (var ifc = new InstalledFontCollection())
            {
                foreach (var family in ifc.Families)
                    cmb.Items.Add(family.Name);
            }
        }

        // ---------------------------------------------------------------
        // Load / Read settings
        // ---------------------------------------------------------------

        private void LoadSettings(DefaultTableSettings settings)
        {
            LoadStyleControls(_header, settings.HeaderStyle);
            LoadStyleControls(_body, settings.BodyStyle);
        }

        private static void LoadStyleControls(StyleControls ctrl, DefaultCellStyle style)
        {
            ctrl.FillColorRGB = style.FillForeColorRGB;
            ctrl.FillColorBtn.BackColor = OfficeBgrToColor(style.FillForeColorRGB);
            ctrl.FillColorBtn.FlatAppearance.MouseOverBackColor = ctrl.FillColorBtn.BackColor;

            ctrl.TransparencyNud.Value = ClampDecimal((decimal)(style.FillTransparency * 100f), 0, 100);

            ctrl.FontColorRGB = style.FontColorRGB;
            ctrl.FontColorBtn.BackColor = OfficeBgrToColor(style.FontColorRGB);
            ctrl.FontColorBtn.FlatAppearance.MouseOverBackColor = ctrl.FontColorBtn.BackColor;

            ctrl.FontNameCmb.Text = style.FontName;
            ctrl.FontSizeNud.Value = ClampDecimal((decimal)style.FontSize, 1, 200);
            ctrl.BoldChk.Checked = style.FontBold;
            ctrl.ItalicChk.Checked = style.FontItalic;

            ctrl.MarginTopNud.Value = ClampDecimal((decimal)style.MarginTop, 0, 100);
            ctrl.MarginBottomNud.Value = ClampDecimal((decimal)style.MarginBottom, 0, 100);
            ctrl.MarginLeftNud.Value = ClampDecimal((decimal)style.MarginLeft, 0, 100);
            ctrl.MarginRightNud.Value = ClampDecimal((decimal)style.MarginRight, 0, 100);

            LoadBorderRow(ctrl.BorderTop, style.BorderTop);
            LoadBorderRow(ctrl.BorderBottom, style.BorderBottom);
            LoadBorderRow(ctrl.BorderLeft, style.BorderLeft);
            LoadBorderRow(ctrl.BorderRight, style.BorderRight);
        }

        private static void LoadBorderRow(BorderRowControls ctrl, DefaultBorderStyle style)
        {
            bool visible = style?.Visible ?? true;
            int colorRGB = style?.ColorRGB ?? 0x000000;
            float weight = style?.Weight ?? 0.75f;
            int dashStyle = style?.DashStyle ?? 1;

            ctrl.VisibleChk.Checked = visible;
            ctrl.ColorRGB = colorRGB;
            ctrl.ColorBtn.BackColor = OfficeBgrToColor(colorRGB);
            ctrl.ColorBtn.FlatAppearance.MouseOverBackColor = ctrl.ColorBtn.BackColor;
            ctrl.WeightNud.Value = ClampDecimal((decimal)weight, 0, 10);
            int idx = dashStyle - 1;
            ctrl.DashCmb.SelectedIndex = (idx >= 0 && idx < ctrl.DashCmb.Items.Count) ? idx : 0;
        }

        private DefaultTableSettings BuildSettings()
        {
            return new DefaultTableSettings
            {
                HeaderStyle = ReadStyleControls(_header),
                BodyStyle = ReadStyleControls(_body)
            };
        }

        private static DefaultCellStyle ReadStyleControls(StyleControls ctrl)
        {
            return new DefaultCellStyle
            {
                FillForeColorRGB = ctrl.FillColorRGB,
                FillTransparency = (float)ctrl.TransparencyNud.Value / 100f,
                FontColorRGB = ctrl.FontColorRGB,
                FontName = ctrl.FontNameCmb.Text,
                FontSize = (float)ctrl.FontSizeNud.Value,
                FontBold = ctrl.BoldChk.Checked,
                FontItalic = ctrl.ItalicChk.Checked,
                MarginTop = (float)ctrl.MarginTopNud.Value,
                MarginBottom = (float)ctrl.MarginBottomNud.Value,
                MarginLeft = (float)ctrl.MarginLeftNud.Value,
                MarginRight = (float)ctrl.MarginRightNud.Value,
                BorderTop = ReadBorderRow(ctrl.BorderTop),
                BorderBottom = ReadBorderRow(ctrl.BorderBottom),
                BorderLeft = ReadBorderRow(ctrl.BorderLeft),
                BorderRight = ReadBorderRow(ctrl.BorderRight)
            };
        }

        private static DefaultBorderStyle ReadBorderRow(BorderRowControls ctrl)
        {
            return new DefaultBorderStyle
            {
                Visible = ctrl.VisibleChk.Checked,
                ColorRGB = ctrl.ColorRGB,
                Weight = (float)ctrl.WeightNud.Value,
                DashStyle = ctrl.DashCmb.SelectedIndex + 1
            };
        }

        // ---------------------------------------------------------------
        // Color conversion helpers (Office BGR ↔ System.Drawing.Color)
        // ---------------------------------------------------------------

        private static Color OfficeBgrToColor(int bgr)
        {
            return Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
        }

        private static int ColorToOfficeBgr(Color c)
        {
            return c.R | (c.G << 8) | (c.B << 16);
        }

        private static decimal ClampDecimal(decimal value, decimal min, decimal max)
        {
            return value < min ? min : value > max ? max : value;
        }

        // ---------------------------------------------------------------
        // Nested helper class to group controls per style section
        // ---------------------------------------------------------------

        private class StyleControls
        {
            public int FillColorRGB { get; set; }
            public int FontColorRGB { get; set; }
            public Button FillColorBtn { get; set; }
            public NumericUpDown TransparencyNud { get; set; }
            public Button FontColorBtn { get; set; }
            public ComboBox FontNameCmb { get; set; }
            public NumericUpDown FontSizeNud { get; set; }
            public CheckBox BoldChk { get; set; }
            public CheckBox ItalicChk { get; set; }
            public NumericUpDown MarginTopNud { get; set; }
            public NumericUpDown MarginBottomNud { get; set; }
            public NumericUpDown MarginLeftNud { get; set; }
            public NumericUpDown MarginRightNud { get; set; }
            public BorderRowControls BorderTop { get; set; }
            public BorderRowControls BorderBottom { get; set; }
            public BorderRowControls BorderLeft { get; set; }
            public BorderRowControls BorderRight { get; set; }
        }

        private class BorderRowControls
        {
            public int ColorRGB { get; set; }
            public CheckBox VisibleChk { get; set; }
            public Button ColorBtn { get; set; }
            public NumericUpDown WeightNud { get; set; }
            public ComboBox DashCmb { get; set; }
        }
    }
}
