using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public sealed class ScatterEditor : Form
    {
        private readonly TextBox title = new TextBox(), xTitle = new TextBox(), yTitle = new TextBox();
        private readonly TextBox xMin = new TextBox(), xMax = new TextBox(), yMin = new TextBox(), yMax = new TextBox();
        private readonly TextBox markerSize = new TextBox(), bubbleSize = new TextBox(), labelSize = new TextBox();
        private readonly CheckBox xLog = new CheckBox(), yLog = new CheckBox(), xReverse = new CheckBox(), yReverse = new CheckBox();
        private readonly CheckBox showGrid = new CheckBox(), showLegend = new CheckBox();
        private readonly ComboBox sizing = new ComboBox(), scheme = new ComboBox(), labels = new ComboBox(), firstContent = new ComboBox(), secondContent = new ComboBox();
        private readonly TextBox workbook = new TextBox(), worksheet = new TextBox(), range = new TextBox();
        private readonly DataGridView data = new DataGridView(), trends = new DataGridView(), partitions = new DataGridView();
        private readonly double width, height;
        private readonly string[] labelKeys = { "None", "Label", "X", "Y", "XY", "Size" };
        private readonly string[] schemeKeys = { "Colors", "Shapes", "None" };
        public ChartSpec Spec { get; private set; }

        public ScatterEditor(ChartSpec initial) : this(initial, 720, 400) { }
        public ScatterEditor(ChartSpec initial, double frameWidth, double frameHeight)
        {
            SuspendLayout();
            Spec = ChartStorage.Clone(initial ?? ChartSamples.Create("Scatter")); ChartEngine.Build(Spec, frameWidth, frameHeight);
            width = frameWidth; height = frameHeight;
            Text = "OpenSlideCharts — Scatter and Bubble"; StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1000, 700); MinimumSize = new Size(800, 580); Font = new Font("Segoe UI", 9);
            BackColor = Color.FromArgb(247, 249, 252);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(18) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); Controls.Add(root);
            var heading = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 26)); heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            heading.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Title — the type follows the Size column: empty = scatter; filled = bubble." }, 0, 0);
            title.Text = Spec.Title ?? ""; title.Dock = DockStyle.Fill; heading.Controls.Add(title, 0, 1); root.Controls.Add(heading, 0, 0);
            var tabs = new TabControl { Dock = DockStyle.Fill }; root.Controls.Add(tabs, 0, 1);
            var dataTab = Tab(tabs, "Data"); var axesTab = Tab(tabs, "Axes"); var markerTab = Tab(tabs, "Markers");
            var labelTab = Tab(tabs, "Labels"); var trendTab = Tab(tabs, "Trend Lines"); var partitionTab = Tab(tabs, "Partitions"); var excelTab = Tab(tabs, "Excel Link");
            ConfigureGrid(data, new string[] { "Label", "X", "Y", "Size", "Group", "Color", "Marker" });
            GridPage(dataTab, data, "Missing X/Y: the point is kept but not plotted. Size ≥ 0; size alone = scale reference. Color #RRGGBB; marker Circle, Square, Triangle, Diamond, Plus, Cross or None.", delegate { PasteClipboard(); });
            data.KeyDown += PasteKey;
            data.EditingControlShowing += delegate(object sender, DataGridViewEditingControlShowingEventArgs e) { e.Control.KeyDown -= PasteKey; e.Control.KeyDown += PasteKey; };
            var axes = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(10) };
            axes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); axes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); axesTab.Controls.Add(axes);
            var xGroup = new GroupBox { Text = "X Axis", Dock = DockStyle.Fill }; var yGroup = new GroupBox { Text = "Y Axis", Dock = DockStyle.Fill };
            axes.Controls.Add(xGroup, 0, 0); axes.Controls.Add(yGroup, 1, 0);
            var xf = Fields(xGroup, 6, 170); var yf = Fields(yGroup, 6, 170);
            Field(xf, 0, "Title", xTitle); Field(xf, 1, "Minimum (empty = auto)", xMin); Field(xf, 2, "Maximum (empty = auto)", xMax);
            Field(xf, 3, "Logarithmic scale", xLog); Field(xf, 4, "Reverse order", xReverse);
            Field(yf, 0, "Title", yTitle); Field(yf, 1, "Minimum (empty = auto)", yMin); Field(yf, 2, "Maximum (empty = auto)", yMax);
            Field(yf, 3, "Logarithmic scale", yLog); Field(yf, 4, "Reverse order", yReverse);
            Note(xf, 5, "Logarithmic: a positive bound is required. Coordinates ≤ 0 are placed at the minimum with a warning. Points outside fixed bounds are not plotted.");
            Note(yf, 5, "Each axis has independent bounds. Date axes and axis breaks are not supported.");
            var markerFields = Fields(markerTab, 7, 260);
            Choices(sizing, new string[] { "Proportional area", "Proportional diameter" }); Choices(scheme, new string[] { "Colors by group", "Shapes by group (scatter)", "No markers" });
            Field(markerFields, 0, "Size representation", sizing); Field(markerFields, 1, "Maximum bubble diameter (pt)", bubbleSize); Field(markerFields, 2, "Marker size (pt)", markerSize);
            Field(markerFields, 3, "Group style", scheme); Field(markerFields, 4, "Gridlines", showGrid); Field(markerFields, 5, "Group legend", showLegend);
            Note(markerFields, 6, "Smaller bubbles are drawn in front of larger ones. Adding a size converts to bubble and restores circular markers. Colors stay tied to groups or points.");
            var labelFields = Fields(labelTab, 5, 260);
            Choices(labels, new string[] { "Automatic (up to 300 points)", "Show", "Hide" });
            string[] labelNames = { "None", "Label", "X value", "Y value", "X and Y", "Size" }; Choices(firstContent, labelNames); Choices(secondContent, labelNames);
            Field(labelFields, 0, "Display", labels); Field(labelFields, 1, "First label", firstContent); Field(labelFields, 2, "Second label", secondContent); Field(labelFields, 3, "Font size (pt)", labelSize);
            Note(labelFields, 4, "Automatic placement from nearby candidate positions; collisions are still possible in dense areas. Long labels may be shortened visually; the data stays complete.");
            ConfigureGrid(trends, new string[] { "Group", "Color" });
            trends.Columns.Insert(0, new DataGridViewCheckBoxColumn { HeaderText = "All Groups", Name = "all" });
            trends.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Show R²", Name = "r2" }); trends.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Equation", Name = "equation" });
            GridPage(trendTab, trends, "Linear trend lines only; linear axes required. All Groups checked = one common regression; otherwise the exact group, empty = ungrouped points. Bubble: weight = raw Size (original choice).", null);
            ConfigureGrid(partitions, new string[] { "Label", "X1", "Y1", "X2", "Y2", "Color" });
            GridPage(partitionTab, partitions, "Lines defined by two distinct points in data coordinates, extended and then clipped to the frame. Both axes must be linear.", null);
            var excel = Fields(excelTab, 5, 260);
            Field(excel, 0, "Saved workbook", workbook); Field(excel, 1, "Sheet", worksheet); Field(excel, 2, "Range without header (3 to 7 columns)", range);
            var excelButtons = new FlowLayoutPanel { Dock = DockStyle.Fill };
            excelButtons.Controls.Add(Button("Browse…", delegate { using (var dialog = new OpenFileDialog { Filter = "Excel workbooks|*.xlsx;*.xlsm;*.xls;*.xlsb", CheckFileExists = true }) if (dialog.ShowDialog(this) == DialogResult.OK) workbook.Text = dialog.FileName; }));
            excelButtons.Controls.Add(Button("Import from Excel", delegate { ImportExcel(); })); excel.Controls.Add(excelButtons, 0, 3); excel.SetColumnSpan(excelButtons, 2);
            Note(excel, 4, "Order: Label, X, Y, Size, Group, Color, Marker. Saved values only. Adding or removing a size converts the type. Macros and external links are not enabled.");
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var apply = Button("Apply", delegate { Apply(); }); var cancel = Button("Cancel", delegate { DialogResult = DialogResult.Cancel; Close(); }); cancel.DialogResult = DialogResult.Cancel;
            footer.Controls.Add(apply); footer.Controls.Add(cancel); root.Controls.Add(footer, 0, 2); AcceptButton = apply; CancelButton = cancel;
            LoadSpec(Spec); ResumeLayout(true);
        }

        private static TabPage Tab(TabControl tabs, string name) { var tab = new TabPage(name); tabs.TabPages.Add(tab); return tab; }
        private static Button Button(string text, EventHandler action) { var button = new Button { Text = text, AutoSize = true, MinimumSize = new Size(110, 28) }; button.Click += action; return button; }
        private static void Choices(ComboBox box, string[] choices) { box.DropDownStyle = ComboBoxStyle.DropDownList; box.Items.AddRange(choices); }
        private static void ConfigureGrid(DataGridView grid, string[] headings)
        {
            grid.Dock = DockStyle.Fill; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; grid.BackgroundColor = Color.White;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize; grid.RowHeadersWidth = 34; grid.AllowUserToOrderColumns = false;
            foreach (string heading in headings) grid.Columns.Add("col" + grid.Columns.Count, heading);
            foreach (DataGridViewColumn column in grid.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
        }
        private static void GridPage(Control parent, DataGridView grid, string text, EventHandler paste)
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38)); parent.Controls.Add(layout);
            layout.Controls.Add(new Label { Text = text, Dock = DockStyle.Fill }, 0, 0); layout.Controls.Add(grid, 0, 1);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill };
            if (paste != null) buttons.Controls.Add(Button("Paste from Excel", paste));
            buttons.Controls.Add(Button("Add Row", delegate { grid.Rows.Add(); }));
            buttons.Controls.Add(Button("Delete Rows", delegate { var rows = new List<DataGridViewRow>(); foreach (DataGridViewCell cell in grid.SelectedCells) if (!cell.OwningRow.IsNewRow && !rows.Contains(cell.OwningRow)) rows.Add(cell.OwningRow); foreach (DataGridViewRow row in rows) grid.Rows.Remove(row); }));
            layout.Controls.Add(buttons, 0, 2);
        }
        private static TableLayoutPanel Fields(Control parent, int rows, int captionWidth)
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = rows, Padding = new Padding(12) };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, captionWidth)); panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < rows; i++) panel.RowStyles.Add(new RowStyle(i == rows - 1 ? SizeType.Percent : SizeType.Absolute, i == rows - 1 ? 100 : 38)); parent.Controls.Add(panel); return panel;
        }
        private static void Field(TableLayoutPanel panel, int row, string text, Control control) { panel.Controls.Add(new Label { Text = text, Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 0) }, 0, row); control.Dock = DockStyle.Fill; panel.Controls.Add(control, 1, row); }
        private static void Note(TableLayoutPanel panel, int row, string text) { var note = new Label { Text = text, Dock = DockStyle.Fill }; panel.Controls.Add(note, 0, row); panel.SetColumnSpan(note, 2); }
        private static string Display(double? value) { return value.HasValue ? ChartEditor.DisplayNumber(value.Value) : ""; }
        private static double? Optional(string value, string location) { return String.IsNullOrWhiteSpace(value) ? (double?)null : ChartEditor.Number(value, location); }
        private static string Cell(DataGridViewRow row, int column) { return Convert.ToString(row.Cells[column].Value, CultureInfo.CurrentCulture) ?? ""; }
        private static bool Flag(DataGridViewRow row, int column) { return row.Cells[column].Value is bool && (bool)row.Cells[column].Value; }
        private static string OptionalText(string value) { return String.IsNullOrWhiteSpace(value) ? null : value.Trim(); }
        private static bool Empty(DataGridViewRow row)
        {
            foreach (DataGridViewCell cell in row.Cells) if (cell.Value is bool ? (bool)cell.Value : !String.IsNullOrWhiteSpace(Convert.ToString(cell.Value, CultureInfo.CurrentCulture))) return false;
            return true;
        }
        private void LoadSpec(ChartSpec spec)
        {
            ScatterSpec options = spec.Scatter;
            ScatterAxisSpec xs = options.XAxis ?? new ScatterAxisSpec(), ys = options.YAxis ?? new ScatterAxisSpec();
            xTitle.Text = options.XTitle ?? ""; yTitle.Text = options.YTitle ?? "";
            xMin.Text = Display(xs.Minimum); xMax.Text = Display(xs.Maximum); yMin.Text = Display(ys.Minimum); yMax.Text = Display(ys.Maximum);
            xLog.Checked = xs.Logarithmic; yLog.Checked = ys.Logarithmic; xReverse.Checked = xs.Reversed; yReverse.Checked = ys.Reversed;
            markerSize.Text = Display(options.MarkerSize); bubbleSize.Text = Display(options.MaxBubbleDiameter); labelSize.Text = Display(options.LabelFontSize);
            sizing.SelectedIndex = options.BubbleSizing == "Diameter" ? 1 : 0; scheme.SelectedIndex = Array.IndexOf(schemeKeys, options.MarkerScheme);
            labels.SelectedIndex = !options.ShowLabels.HasValue ? 0 : options.ShowLabels.Value ? 1 : 2;
            firstContent.SelectedIndex = Array.IndexOf(labelKeys, options.LabelContent); secondContent.SelectedIndex = Array.IndexOf(labelKeys, options.SecondLabelContent);
            showGrid.Checked = options.ShowGrid; showLegend.Checked = options.ShowLegend;
            workbook.Text = spec.WorkbookPath ?? ""; worksheet.Text = spec.Worksheet ?? ""; range.Text = spec.RangeAddress ?? "";
            LoadPoints(options.Points);
            if (options.Trendlines != null) foreach (ScatterTrendline trend in options.Trendlines) trends.Rows.Add(trend.Group == null, trend.Group ?? "", trend.Color ?? "", trend.ShowRSquared, trend.ShowEquation);
            if (options.Partitions != null) foreach (ScatterPartition partition in options.Partitions) partitions.Rows.Add(partition.Label ?? "", Display(partition.X1), Display(partition.Y1), Display(partition.X2), Display(partition.Y2), partition.Color ?? "");
        }
        private void LoadPoints(List<ScatterPoint> points)
        {
            data.CancelEdit(); data.Rows.Clear();
            foreach (ScatterPoint point in points) data.Rows.Add(point.Label ?? "", Display(point.X), Display(point.Y), Display(point.Size), point.Group ?? "", point.Color ?? "", point.Marker ?? "");
        }
        private ChartSpec ReadSpec(bool readPoints)
        {
            data.EndEdit(); trends.EndEdit(); partitions.EndEdit(); ChartSpec result = ChartStorage.Clone(Spec); ScatterSpec options = result.Scatter;
            result.Title = title.Text.Trim(); result.WorkbookPath = workbook.Text.Trim(); result.Worksheet = worksheet.Text.Trim(); result.RangeAddress = range.Text.Trim();
            options.XTitle = xTitle.Text.Trim(); options.YTitle = yTitle.Text.Trim();
            options.XAxis = new ScatterAxisSpec { Minimum = Optional(xMin.Text, "Minimum X"), Maximum = Optional(xMax.Text, "Maximum X"), Logarithmic = xLog.Checked, Reversed = xReverse.Checked };
            options.YAxis = new ScatterAxisSpec { Minimum = Optional(yMin.Text, "Minimum Y"), Maximum = Optional(yMax.Text, "Maximum Y"), Logarithmic = yLog.Checked, Reversed = yReverse.Checked };
            options.MarkerSize = ChartEditor.Number(markerSize.Text, "Marker size"); options.MaxBubbleDiameter = ChartEditor.Number(bubbleSize.Text, "Bubble diameter"); options.LabelFontSize = ChartEditor.Number(labelSize.Text, "Label font size");
            options.BubbleSizing = sizing.SelectedIndex == 1 ? "Diameter" : "Area"; options.MarkerScheme = schemeKeys[scheme.SelectedIndex];
            options.ShowLabels = labels.SelectedIndex == 0 ? (bool?)null : labels.SelectedIndex == 1; options.LabelContent = labelKeys[firstContent.SelectedIndex]; options.SecondLabelContent = labelKeys[secondContent.SelectedIndex];
            options.ShowGrid = showGrid.Checked; options.ShowLegend = showLegend.Checked; options.Trendlines = new List<ScatterTrendline>(); options.Partitions = new List<ScatterPartition>();
            foreach (DataGridViewRow row in trends.Rows) if (!row.IsNewRow) options.Trendlines.Add(new ScatterTrendline { Group = Flag(row, 0) ? null : Cell(row, 1), Color = OptionalText(Cell(row, 2)), ShowRSquared = Flag(row, 3), ShowEquation = Flag(row, 4) });
            foreach (DataGridViewRow row in partitions.Rows) if (!row.IsNewRow && !Empty(row)) options.Partitions.Add(new ScatterPartition { Label = Cell(row, 0), X1 = ChartEditor.Number(Cell(row, 1), "Partition X1"), Y1 = ChartEditor.Number(Cell(row, 2), "Partition Y1"), X2 = ChartEditor.Number(Cell(row, 3), "Partition X2"), Y2 = ChartEditor.Number(Cell(row, 4), "Partition Y2"), Color = OptionalText(Cell(row, 5)) });
            if (!readPoints) { options.Points = new List<ScatterPoint>(); return result; }
            var points = new List<ScatterPoint>();
            foreach (DataGridViewRow row in data.Rows) if (!row.IsNewRow && !Empty(row)) points.Add(new ScatterPoint { Label = Cell(row, 0), X = Optional(Cell(row, 1), "X row " + (row.Index + 1)), Y = Optional(Cell(row, 2), "Y row " + (row.Index + 1)), Size = Optional(Cell(row, 3), "Size row " + (row.Index + 1)), Group = Cell(row, 4), Color = OptionalText(Cell(row, 5)), Marker = OptionalText(Cell(row, 6)) });
            return ScatterData.ReplacePoints(result, points);
        }
        private void ValidateCandidate(ChartSpec candidate) { ChartEngine.Build(candidate, width, height); }
        private void Apply()
        {
            try { ChartSpec candidate = ReadSpec(true); ValidateCandidate(candidate); Spec = candidate; DialogResult = DialogResult.OK; Close(); }
            catch (Exception error) { ShowError(error); }
        }
        private void PasteKey(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.KeyCode != Keys.V) return;
            try { string text = Clipboard.GetText(); if (text.IndexOf('\t') < 0 && text.TrimEnd('\r', '\n').IndexOf('\n') < 0) return; e.Handled = true; e.SuppressKeyPress = true; PasteData(text); }
            catch (Exception error) { ShowError(error); }
        }
        private void PasteClipboard() { try { PasteData(Clipboard.GetText()); } catch (Exception error) { ShowError(error); } }
        private void PasteData(string text)
        {
            var points = new List<ScatterPoint>(); var table = ChartEditor.ParseTsv(text);
            foreach (List<string> cells in table)
            {
                if (cells.TrueForAll(delegate(string cell) { return String.IsNullOrWhiteSpace(cell); })) continue;
                if (points.Count == 0 && cells.Count >= 3 && cells[1].Trim().Equals("X", StringComparison.OrdinalIgnoreCase) && cells[2].Trim().Equals("Y", StringComparison.OrdinalIgnoreCase)) continue;
                if (cells.Count < 3 || cells.Count > 7) throw new ArgumentException("Paste: 3 to 7 columns expected (Label, X, Y, Size, Group, Color, Marker).");
                points.Add(new ScatterPoint { Label = cells[0], X = Optional(cells[1], "Pasted X"), Y = Optional(cells[2], "Pasted Y"), Size = cells.Count > 3 ? Optional(cells[3], "Pasted size") : null,
                    Group = cells.Count > 4 ? cells[4] : "", Color = cells.Count > 5 ? OptionalText(cells[5]) : null, Marker = cells.Count > 6 ? OptionalText(cells[6]) : null });
            }
            ChartSpec candidate = ScatterData.ReplacePoints(ReadSpec(false), points); ValidateCandidate(candidate);
            LoadPoints(candidate.Scatter.Points); scheme.SelectedIndex = Array.IndexOf(schemeKeys, candidate.Scatter.MarkerScheme);
        }
        private void ImportExcel()
        {
            try { UseWaitCursor = true; ChartSpec source = ReadExcelSource(); ChartSpec candidate = ExcelLink.Read(source); ValidateCandidate(candidate); LoadPoints(candidate.Scatter.Points); scheme.SelectedIndex = Array.IndexOf(schemeKeys, candidate.Scatter.MarkerScheme); }
            catch (Exception error) { ShowError(error); }
            finally { UseWaitCursor = false; }
        }
        private ChartSpec ReadExcelSource()
        {
            ChartSpec source = ReadSpec(false);
            // Preserve positional styles while allowing Excel to replace invalid
            // numeric text currently being edited in the data grid.
            foreach (DataGridViewRow row in data.Rows) if (!row.IsNewRow && !Empty(row))
                source.Scatter.Points.Add(new ScatterPoint { Label = Cell(row, 0), Group = Cell(row, 4), Color = OptionalText(Cell(row, 5)), Marker = OptionalText(Cell(row, 6)) });
            return source;
        }
        private void ShowError(Exception error) { MessageBox.Show(this, error.Message, "Check Data", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
}
