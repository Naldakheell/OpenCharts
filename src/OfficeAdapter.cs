using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;

namespace OpenSlideCharts
{
    public static class ChartStorage
    {
        private const int MaximumEncodedCharacters = 2 * 1024 * 1024;
        private const int TagCharacters = 180;
        private const int MaximumTags = (MaximumEncodedCharacters + TagCharacters - 1) / TagCharacters;
        public static string Serialize(ChartSpec spec) { return new JavaScriptSerializer().Serialize(spec); }
        public static ChartSpec Clone(ChartSpec spec) { return new JavaScriptSerializer().Deserialize<ChartSpec>(Serialize(spec)); }
        public static void Write(object shapeObject, ChartSpec spec)
        {
            dynamic shape = shapeObject;
            string hex = BitConverter.ToString(Encoding.UTF8.GetBytes(Serialize(spec))).Replace("-", "");
            if (hex.Length > MaximumEncodedCharacters) throw new ArgumentException("Data too large: 1 MiB of UTF-8 JSON maximum.");
            int count = (hex.Length + TagCharacters - 1) / TagCharacters;
            shape.Tags.Add("OSC_VERSION", "1");
            shape.Tags.Add("OSC_COUNT", count.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < count; i++) shape.Tags.Add("OSC_DATA_" + i.ToString("D4"), hex.Substring(i * TagCharacters, Math.Min(TagCharacters, hex.Length - i * TagCharacters)));
            shape.AlternativeText = "OpenSlideCharts: " + spec.Title + ". " + spec.Kind + " chart.";
        }
        public static ChartSpec Read(object shapeObject)
        {
            dynamic shape = shapeObject;
            string version = shape.Tags.Item("OSC_VERSION");
            if (version != "1") throw new ArgumentException("Select an OpenSlideCharts group (version 1).");
            int count;
            if (!int.TryParse((string)shape.Tags.Item("OSC_COUNT"), out count) || count < 1 || count > MaximumTags) throw new InvalidDataException("Invalid chart metadata.");
            StringBuilder encoded = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                string part = shape.Tags.Item("OSC_DATA_" + i.ToString("D4"));
                if (part.Length == 0 || part.Length > TagCharacters) throw new InvalidDataException("Incomplete chart data.");
                encoded.Append(part);
            }
            string hex = encoded.ToString();
            if (hex.Length > MaximumEncodedCharacters) throw new InvalidDataException("Chart data too large.");
            if (hex.Length % 2 != 0) throw new InvalidDataException("Invalid chart encoding.");
            byte[] bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++) bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
            ChartSpec spec = new JavaScriptSerializer().Deserialize<ChartSpec>(Encoding.UTF8.GetString(bytes));
            if (spec == null) throw new InvalidDataException("Empty chart.");
            return spec;
        }
    }

    public static class OfficeRenderer
    {
        public static ChartRectangle InsertionFrame(ChartSpec spec, double slideWidth, double slideHeight)
        {
            ChartRectangle frame = spec.Appearance == null ? null : spec.Appearance.SourceFrame;
            if (frame == null) return new ChartRectangle { X = 30, Y = 35, Width = Math.Max(320, slideWidth - 60), Height = Math.Max(240, slideHeight - 70) };
            spec.Appearance.Validate(spec.Series.Count, spec.Rows.Count);
            if (frame.X < 0 || frame.Y < 0 || frame.X + frame.Width > slideWidth + .001 || frame.Y + frame.Height > slideHeight + .001)
                throw new ArgumentException("The original frame extends beyond this slide. Use a slide with the same dimensions as the source.");
            return frame;
        }
        // Render first, then replace. A failed render leaves the previous chart intact.
        public static object Render(object slideObject, ChartSpec spec, double left, double top, double width, double height)
        {
            Layout layout = ChartEngine.Build(spec, width, height);
            dynamic slide = slideObject;
            List<string> names = new List<string>();
            string prefix = "OSC_" + Guid.NewGuid().ToString("N") + "_";
            dynamic group = null;
            try
            {
                dynamic bounds = slide.Shapes.AddShape(1, (float)left, (float)top, (float)width, (float)height);
                bounds.Name = prefix + "bounds";
                names.Add((string)bounds.Name);
                bounds.Fill.Visible = -1;
                bounds.Fill.Solid();
                bounds.Fill.Transparency = 1f;
                bounds.Line.Visible = 0;
                for (int itemIndex = 0; itemIndex < layout.Items.Count; itemIndex++)
                {
                    Primitive p = layout.Items[itemIndex];
                    dynamic s;
                    if (p.Kind == "line")
                    {
                        if (p.Role == "WATERFALL_CONNECTOR")
                            s = slide.Shapes.AddConnector(1, (float)(left + p.X), (float)(top + p.Y), (float)(left + p.X + p.W), (float)(top + p.Y + p.H));
                        else s = slide.Shapes.AddLine((float)(left + p.X), (float)(top + p.Y), (float)(left + p.X + p.W), (float)(top + p.Y + p.H));
                        s.Line.ForeColor.RGB = p.Color;
                        s.Line.Weight = (float)(p.LineWidth > 0 ? p.LineWidth : 1.3);
                        if (p.DashStyle > 0) s.Line.DashStyle = p.DashStyle;
                    }
                    else if (p.Kind == "text")
                    {
                        s = slide.Shapes.AddTextbox(1, (float)(left + p.X), (float)(top + p.Y), (float)Math.Max(1, p.W), (float)Math.Max(1, p.H));
                        s.TextFrame.AutoSize = 0;
                        s.TextFrame.MarginLeft = (float)(p.MarginLeft ?? 1); s.TextFrame.MarginRight = (float)(p.MarginRight ?? 1);
                        s.TextFrame.MarginTop = 0f; s.TextFrame.MarginBottom = 0f;
                        s.TextFrame.WordWrap = p.NoWrap ? 0 : -1;
                        if (p.VerticalAlignment != null) s.TextFrame.VerticalAnchor = p.VerticalAlignment == "bottom" ? 4 : (p.VerticalAlignment == "center" ? 3 : 1);
                        s.TextFrame.AutoSize = 0;
                        s.TextFrame.TextRange.Text = p.Text ?? "";
                        s.TextFrame.TextRange.Font.Name = String.IsNullOrEmpty(p.FontName) ? "Aptos" : p.FontName;
                        s.TextFrame.TextRange.Font.Size = (float)(p.FontSize > 0 ? p.FontSize : 11);
                        s.TextFrame.TextRange.Font.Color.RGB = p.Color;
                        s.TextFrame.TextRange.ParagraphFormat.Alignment = p.Alignment == "right" ? 3 : (p.Alignment == "center" ? 2 : 1);
                        // Office may resize an empty textbox while changing wrapping or anchors.
                        // Restore the requested geometry after configuring its text frame.
                        s.Left = (float)(left + p.X); s.Top = (float)(top + p.Y);
                        s.Width = (float)Math.Max(1, p.W); s.Height = (float)Math.Max(1, p.H);
                    }
                    else if (p.Kind == "path")
                    {
                        s = RenderPath((object)slide, p, left, top);
                    }
                    else if (p.Kind == "rect" || p.Kind == "ellipse")
                    {
                        bool exactSize = p.Role == "SCATTER_POINT";
                        s = slide.Shapes.AddShape(p.Kind == "ellipse" ? 9 : 1, (float)(left + p.X), (float)(top + p.Y), (float)(exactSize ? p.W : Math.Max(.05, p.W)), (float)(exactSize ? p.H : Math.Max(.05, p.H)));
                        s.Fill.Solid(); s.Fill.ForeColor.RGB = p.Color; s.Line.Visible = 0;
                    }
                    else throw new NotSupportedException("Unknown primitive: " + p.Kind);
                    s.Name = prefix + names.Count.ToString(CultureInfo.InvariantCulture);
                    if (!String.IsNullOrEmpty(p.Role))
                    {
                        s.Tags.Add("OSC_ITEM", itemIndex.ToString(CultureInfo.InvariantCulture));
                        s.Tags.Add("OSC_ROLE", p.Role);
                    }
                    names.Add((string)s.Name);
                }
                object[] indices = names.ConvertAll<object>(delegate(string n) { return n; }).ToArray();
                group = slide.Shapes.Range(indices).Group();
                group.Name = "OpenSlideCharts_" + Guid.NewGuid().ToString("N");
                ChartStorage.Write((object)group, spec);
                return group;
            }
            catch
            {
                if (group != null) { try { group.Delete(); } catch { } }
                else foreach (string name in names) { try { slide.Shapes.Item(name).Delete(); } catch { } }
                throw;
            }
        }
        private static object RenderPath(object slideObject, Primitive primitive, double left, double top)
        {
            ShapePath path = primitive.Path;
            if (path == null || path.Segments == null || path.Segments.Count == 0)
                throw new ArgumentException("A closed outline is required for a freeform shape.");
            dynamic slide = slideObject;
            // Office's freeform consists of one closed contour. An annular sector
            // follows the outer arc, then the inner arc in reverse; its hole is
            // genuinely unfilled, including the coincident seam of a full ring.
            dynamic builder = slide.Shapes.BuildFreeform(1, (float)(left + path.StartX), (float)(top + path.StartY));
            foreach (PathSegment segment in path.Segments)
            {
                if (segment.Curve)
                    builder.AddNodes(1, 1, (float)(left + segment.C1X), (float)(top + segment.C1Y),
                        (float)(left + segment.C2X), (float)(top + segment.C2Y), (float)(left + segment.X), (float)(top + segment.Y));
                else builder.AddNodes(0, 1, (float)(left + segment.X), (float)(top + segment.Y));
            }
            dynamic shape = builder.ConvertToShape();
            try
            {
                shape.Fill.Solid(); shape.Fill.ForeColor.RGB = primitive.Color; shape.Line.Visible = 0;
                return (object)shape;
            }
            catch { try { shape.Delete(); } catch { } throw; }
        }
        public static object Replace(object slideObject, object oldObject, ChartSpec spec)
        { return Replace(slideObject, oldObject, spec, null); }
        public static object Replace(object slideObject, object oldObject, ChartSpec spec, ChartRectangle frame)
        {
            dynamic old = oldObject;
            // Rotated groups cannot be regenerated without changing their local axes yet.
            if (Math.Abs((double)old.Rotation) > .01) throw new NotSupportedException("Reset the group rotation to 0° before refreshing it.");
            int position = old.ZOrderPosition;
            dynamic fresh = Render(slideObject, spec, frame == null ? (double)old.Left : frame.X, frame == null ? (double)old.Top : frame.Y, frame == null ? (double)old.Width : frame.Width, frame == null ? (double)old.Height : frame.Height);
            try
            {
                while ((int)fresh.ZOrderPosition > position) fresh.ZOrder(3);
                old.Delete();
                return fresh;
            }
            catch { try { fresh.Delete(); } catch { } throw; }
        }
    }

    public static class ExcelLink
    {
        public static double Number(object value)
        {
            if (value is double || value is int || value is float || value is decimal || value is long) return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            string text = Convert.ToString(value, CultureInfo.CurrentCulture).Trim().Replace("\u00a0", "").Replace("\u202f", "");
            double result;
            // No thousands separator: avoid silently interpreting 1,234 as 1234.
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out result) || double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result)) return result;
            throw new FormatException("Invalid number: " + text);
        }
        public static ChartSpec Read(ChartSpec source)
        {
            if (source != null && source.Waterfall != null) throw new NotSupportedException("Excel links are not yet supported for waterfalls with connectors. Paste the data in the waterfall editor instead.");
            if (source != null && source.Series != null && source.Series.Count > 0) throw new NotSupportedException("Multi-series Excel links are not yet supported. Paste a matrix in the editor instead.");
            if (source == null || String.IsNullOrWhiteSpace(source.WorkbookPath) || String.IsNullOrWhiteSpace(source.Worksheet) || String.IsNullOrWhiteSpace(source.RangeAddress)) throw new ArgumentException("Specify the Excel file, sheet and range without a header (for example, A2:D8).");
            if (!File.Exists(source.WorkbookPath)) throw new FileNotFoundException("Workbook not found; the embedded data is kept.", source.WorkbookPath);
            dynamic app = null, books = null, book = null, sheets = null, sheet = null, range = null, rows = null, columns = null, areas = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
                app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false;
                app.AutomationSecurity = 3; app.AskToUpdateLinks = false;
                books = app.Workbooks;
                book = books.Open(Path.GetFullPath(source.WorkbookPath), 0, true, Type.Missing, "", "", true, Type.Missing, Type.Missing, false, false, Type.Missing, false);
                sheets = book.Worksheets; sheet = sheets.Item(source.Worksheet); range = sheet.Range[source.RangeAddress];
                rows = range.Rows; columns = range.Columns; areas = range.Areas;
                int rowCount = rows.Count, colCount = columns.Count;
                if (source.Gantt != null)
                {
                    if (rowCount < 1 || rowCount > 101 || colCount < 2 || colCount > 33 || (int)areas.Count != 1)
                        throw new ArgumentException("Gantt: contiguous range of 1 to 101 rows and 2 to 33 columns; the first row is reserved for milestones.");
                    return GanttData.ReadCells(source, range.Value2 as object[,], (bool)book.Date1904);
                }
                if (String.Equals(source.Kind, "Scatter", StringComparison.OrdinalIgnoreCase) || String.Equals(source.Kind, "Bubble", StringComparison.OrdinalIgnoreCase))
                {
                    if (rowCount < 1 || rowCount > 1000 || colCount < 3 || colCount > 7 || (int)areas.Count != 1)
                        throw new ArgumentException("Scatter/bubble: contiguous range of 1 to 1000 rows and 3 to 7 columns, without a header.");
                    return ScatterData.ReadCells(source, range.Value2 as object[,]);
                }
                if (rowCount < 1 || rowCount > 40 || colCount < 2 || colCount > 4 || (int)areas.Count != 1) throw new ArgumentException("Contiguous range expected: 1 to 40 rows, 2 to 4 columns, without a header.");
                object[,] cells = range.Value2 as object[,];
                if (cells == null) throw new ArgumentException("The range must contain a label and a value.");
                ChartSpec result = ChartStorage.Clone(source); result.Rows.Clear();
                for (int row = 1; row <= rowCount; row++)
                {
                    string label = Convert.ToString(cells[row, 1], CultureInfo.CurrentCulture);
                    if (String.IsNullOrWhiteSpace(label) && cells[row, 2] == null) continue;
                    if (cells[row, 2] == null) throw new FormatException("Missing value in row " + row);
                    ChartRow item = new ChartRow { Label = label, Value = Number(cells[row, 2]) };
                    item.Secondary = colCount > 2 && cells[row, 3] != null ? Number(cells[row, 3]) : 0;
                    if (colCount > 3 && cells[row, 4] != null)
                    {
                        string flag = Convert.ToString(cells[row, 4], CultureInfo.InvariantCulture).Trim().ToLowerInvariant();
                        if (flag == "true" || flag == "vrai" || flag == "1" || flag == "oui") item.Total = true;
                        else if (flag != "false" && flag != "faux" && flag != "0" && flag != "non" && flag != "") throw new FormatException("Invalid Total flag in row " + row);
                    }
                    result.Rows.Add(item);
                }
                if (result.Pie != null && result.Pie.Slices != null && result.Pie.Slices.Count > 0 && source.Rows != null && source.Rows.Count != result.Rows.Count)
                {
                    // Row indexes identify slice styles. A changed row count does
                    // not establish which source rows were inserted or removed.
                    result.Pie.Slices.Clear();
                    if (result.ImportWarnings == null) result.ImportWarnings = new List<string>();
                    const string warning = "Le nombre de lignes du secteur/anneau a changé : couleurs et éclatements individuels réinitialisés.";
                    if (!result.ImportWarnings.Contains(warning)) result.ImportWarnings.Add(warning);
                }
                if (String.Equals(result.Kind, "Pie", StringComparison.OrdinalIgnoreCase) || String.Equals(result.Kind, "Doughnut", StringComparison.OrdinalIgnoreCase))
                    ChartEngine.ValidatePie(result);
                else ChartEngine.Build(result, 720, 400);
                return result;
            }
            finally
            {
                Release((object)areas); Release((object)columns); Release((object)rows);
                Release((object)range); Release((object)sheet); Release((object)sheets);
                if (book != null) { try { book.Close(false); } catch { } }
                Release((object)book); Release((object)books);
                if (app != null) { try { app.Quit(); } catch { } }
                Release((object)app);
            }
        }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
    }
}
