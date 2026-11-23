using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public sealed partial class WaterfallSlideEditor
    {
        private WaterfallTarget pendingSource;
        public bool AddingConnector { get { return Active && pendingSource != null; } }
        private sealed class CommandMenu : ContextMenuStrip
        {
            internal Action EscapeAction;
            protected override bool ProcessDialogKey(Keys keyData)
            {
                if ((keyData & Keys.KeyCode) == Keys.Escape && EscapeAction != null) { EscapeAction(); return true; }
                return base.ProcessDialogKey(keyData);
            }
        }
        private sealed class CommandSubmenu : ToolStripDropDownMenu
        {
            internal Action EscapeAction;
            protected override bool ProcessDialogKey(Keys keyData)
            {
                if ((keyData & Keys.KeyCode) == Keys.Escape && EscapeAction != null) { EscapeAction(); return true; }
                return base.ProcessDialogKey(keyData);
            }
        }

        public void BeginSegmentConnector(object slideObject, object chartObject, string columnId, string segmentId, string edge)
        {
            StartForSegment(slideObject, chartObject);
            try
            {
                pendingSource = SegmentTarget(columnId, segmentId, edge);
                SelectSegment(columnId, segmentId);
                LastMessage = "Select the segment where the connector ends.";
            }
            catch { Stop(false); throw; }
        }
        public void CancelConnectorCreation() { Stop(false, true); }
        public void CompleteSegmentConnector(object chartObject, string columnId, string segmentId, string edge, bool constrain = true)
        {
            RequireCommandChart();
            if (!SameObject(chartObject, chart)) throw new InvalidOperationException("Choose a destination in the same chart.");
            if (!AddingConnector) throw new InvalidOperationException("Choose the connector's starting point first.");
            WaterfallTarget source = pendingSource, target = SegmentTarget(columnId, segmentId, edge);
            if (WaterfallInteraction.SameAnchor(source.Anchor, target.Anchor)) throw new ArgumentException("Choose two different levels.");
            WaterfallCommandAnalysis analysis = WaterfallCommands.AnalyzeAddConnector(spec.Waterfall, source.Anchor, target.Anchor, constrain);
            ExecuteWithRepair(analysis, delegate(string release)
            {
                CloseCommandMenu();
                RequireCommandChart();
                ChartSpec candidate = ChartStorage.Clone(spec);
                WaterfallCommandResult result = WaterfallCommands.AddConnector(candidate.Waterfall, source.Anchor, target.Anchor, constrain, release);
                string previousId = currentId;
                pendingSource = null; currentId = result.ConnectorId;
                bool before = busy; busy = true;
                try { Apply(candidate); LastMessage = "Connector added; e values recalculated."; }
                catch { pendingSource = source; currentId = previousId; throw; }
                finally { busy = before; }
            });
        }
        public void PinSegmentZero(object slideObject, object chartObject, string columnId, string segmentId, string edge)
        {
            StartForSegment(slideObject, chartObject);
            WaterfallTarget target = SegmentTarget(columnId, segmentId, edge);
            WaterfallCommandAnalysis analysis = WaterfallCommands.AnalyzePinZero(spec.Waterfall, target.Anchor);
            ExecuteWithRepair(analysis, delegate(string release)
            {
                CloseCommandMenu();
                RequireCommandChart();
                ChartSpec candidate = ChartStorage.Clone(spec);
                WaterfallCommands.PinZero(candidate.Waterfall, target.Anchor, release);
                bool before = busy, beforeSuppress = suppressSelection;
                busy = true; suppressSelection = true;
                try { Apply(candidate); SelectSegment(columnId, segmentId); LastMessage = "Segment edge anchored at zero."; }
                finally { busy = before; suppressSelection = beforeSuppress; }
            });
        }
        private void StartForSegment(object slideObject, object chartObject)
        {
            bool before = suppressSelection; suppressSelection = true;
            try { Start(slideObject, chartObject); }
            finally { suppressSelection = before; }
        }
        private WaterfallTarget SegmentTarget(string columnId, string segmentId, string edge)
        {
            RequireCommandChart();
            WaterfallTarget result = scene.Targets.Find(delegate(WaterfallTarget t)
            { return t.ColumnId == columnId && t.Edge == edge && t.Side == "Left" && ((edge == "Base" || edge == "Output") || t.SegmentId == segmentId); });
            if (result == null) throw new ArgumentException("This segment level no longer exists.");
            return result;
        }
        private void SelectSegment(string columnId, string segmentId)
        {
            bool before = busy; busy = true;
            try
            {
                Layout layout = ChartEngine.Build(spec, width, height); dynamic group = chart;
                for (int i = 1; i <= (int)group.GroupItems.Count; i++)
                {
                    dynamic shape = group.GroupItems.Item(i); int index;
                    if (!Int32.TryParse((string)shape.Tags.Item("OSC_ITEM"), out index) || index < 0 || index >= layout.Items.Count) continue;
                    Primitive item = layout.Items[index];
                    if (item.Role == "WATERFALL_SEGMENT" && item.ColumnId == columnId && item.SegmentId == segmentId) { shape.Select(); return; }
                }
                throw new InvalidOperationException("The native segment is no longer available.");
            }
            finally { busy = before; }
        }
        private void RequireCommandChart()
        {
            if (!Active || !ChartExists() || !IsActiveSlide() || !FrameUnchanged() || !Upright(chart))
                throw new InvalidOperationException("The selection or frame changed. Run the command again.");
        }
        private void PollConnectorCreation()
        {
            if (!Active || busy) return;
            if (!ChartExists() || !IsActiveSlide() || !FrameUnchanged() || !Upright(chart) || (GetAsyncKeyState(27) & 0x8000) != 0)
            { bool before = suppressSelection; suppressSelection = true; try { Stop(false); } finally { suppressSelection = before; } return; }
            if (!choosing)
            {
                OfficeChartSelection selected = OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection);
                if (selected == null || !SameObject(selected.Chart, chart))
                { bool before = suppressSelection; suppressSelection = true; try { Stop(false); } finally { suppressSelection = before; } }
            }
        }
        private void CreationSelectionChanged(object selectionObject)
        {
            OfficeChartSelection selected = OfficeChartSelection.Resolve(selectionObject);
            if (selected == null || !SameObject(selected.Chart, chart))
            { bool before = suppressSelection; suppressSelection = true; try { Stop(false); } finally { suppressSelection = before; } return; }
            if (selected.Primitive == null || selected.Primitive.Role != "WATERFALL_SEGMENT") return;
            Primitive destination = selected.Primitive; object destinationChart = selected.Chart; string session = token;
            var choices = new CommandMenu();
            AddDestination(choices, destinationChart, session, destination, "Top", "Segment Top");
            AddDestination(choices, destinationChart, session, destination, "Bottom", "Segment Bottom");
            AddDestination(choices, destinationChart, session, destination, "Output", "Column Cumulative");
            choices.Items.Add("Cancel Adding Connector", null, delegate { Run(delegate { CancelConnectorCreation(); }); });
            ShowCommandMenu(choices);
            LastMessage = "Choose the end level and the connection type.";
        }
        private void AddDestination(ContextMenuStrip choices, object destinationChart, string session, Primitive destination, string edge, string label)
        {
            var entry = new ToolStripMenuItem(label); entry.DropDown = new CommandSubmenu();
            entry.DropDownItems.Add("Equalize Levels", null, delegate { if (!Active || token != session) return; Run(delegate { choosing = false; CompleteSegmentConnector(destinationChart, destination.ColumnId, destination.SegmentId, edge, true); }); });
            entry.DropDownItems.Add("Free Connector", null, delegate { if (!Active || token != session) return; Run(delegate { choosing = false; CompleteSegmentConnector(destinationChart, destination.ColumnId, destination.SegmentId, edge, false); }); });
            choices.Items.Add(entry);
        }
        private void ExecuteWithRepair(WaterfallCommandAnalysis analysis, Action<string> commit)
        {
            if (analysis.StrictSucceeds) { commit(null); return; }
            if (analysis.AutoPinIds.Count == 1) { commit(analysis.AutoPinIds[0]); return; }
            if (analysis.AutoPinIds.Count == 0)
                throw new InvalidOperationException(analysis.StrictFailure == WaterfallSolveFailure.Inconsistent
                    ? "Conflicting anchors: change the anchors or choose a free connector."
                    : "This command requires an additional anchor in the data editor.");
            var choices = new CommandMenu(); string session = token;
            foreach (string item in analysis.AutoPinIds)
            {
                string id = item; WaterfallPin pin = spec.Waterfall.Pins.Find(delegate(WaterfallPin p) { return p.Id == id; });
                choices.Items.Add("Move Automatic Anchor: " + AnchorLabel(pin.Anchor), null, delegate
                {
                    if (!Active || token != session) return;
                    Run(delegate { choosing = false; commit(id); });
                });
            }
            choices.Items.Add("Keep Anchors");
            ShowCommandMenu(choices);
            LastMessage = "Several automatic anchors can be moved. Choose one.";
        }
        private string AnchorLabel(WaterfallAnchor anchor)
        {
            WaterfallTarget target = scene.Targets.Find(delegate(WaterfallTarget t) { return WaterfallInteraction.SameAnchor(t.Anchor, anchor); });
            if (target != null) return TargetLabel(target);
            WaterfallColumn column = spec.Waterfall.Columns.Find(delegate(WaterfallColumn c) { return c.Id == anchor.ColumnId; });
            return column.Label + " — linked level";
        }
        private void ShowCommandMenu(ContextMenuStrip choices)
        {
            CloseCommandMenu();
            menu = choices; choosing = true;
            BindCommandEscape(choices);
            choices.Closed += delegate(object sender, ToolStripDropDownClosedEventArgs args)
            {
                if (!Object.ReferenceEquals(menu, choices)) return;
                menu = null; choosing = false;
                if (args.CloseReason == ToolStripDropDownCloseReason.Keyboard && AddingConnector) Run(CancelConnectorCreation);
                DisposeMenuLater(choices);
            };
            choices.Show(Cursor.Position);
        }
        private void BindCommandEscape(ToolStripDropDown dropdown)
        {
            Action escape = delegate { Run(delegate { if (AddingConnector) CancelConnectorCreation(); else CloseCommandMenu(); }); };
            var root = dropdown as CommandMenu; if (root != null) root.EscapeAction = escape;
            var child = dropdown as CommandSubmenu; if (child != null) child.EscapeAction = escape;
            foreach (ToolStripItem item in dropdown.Items)
            {
                var entry = item as ToolStripDropDownItem;
                if (entry != null && entry.HasDropDownItems) BindCommandEscape(entry.DropDown);
            }
        }
        private void CloseCommandMenu()
        {
            ContextMenuStrip previous = menu; menu = null; choosing = false;
            if (previous != null) { previous.Close(); DisposeMenuLater(previous); }
        }
        private static void DisposeMenuLater(ContextMenuStrip choices)
        {
            // Closed precedes both SetVisibleCore's cleanup and the item's Click.
            // Disposing in that callback can make WinForms recreate a disposed handle.
            if (choices.IsDisposed) return;
            if (choices.IsHandleCreated) choices.BeginInvoke((MethodInvoker)delegate { choices.Dispose(); });
            else choices.Dispose();
        }
    }
}
