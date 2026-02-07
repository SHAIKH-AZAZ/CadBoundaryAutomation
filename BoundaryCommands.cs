using System;
using System.Collections.Generic;

// WinForms / Drawing aliases to avoid AutoCAD name conflicts
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(CadBoundaryAutomation.BoundaryCommands))]

namespace CadBoundaryAutomation
{
    public class BoundaryCommands
    {
        private enum BarsOrientation
        {
            Horizontal,
            Vertical,
            Both
        }

        private class SessionState
        {
            public Document Doc;
            public Editor Ed;
            public Database Db;

            public double Spacing;
            public BarsOrientation Orientation;

            public ObjectId CreatedBoundaryId = ObjectId.Null;

            public ObjectEventHandler AppendedHandler;
            public CommandEventHandler EndedHandler;
            public CommandEventHandler CancelledHandler;
            public CommandEventHandler FailedHandler;
            public EventHandler IdleHandler;
        }

        private static readonly Dictionary<Document, SessionState> _sessions =
            new Dictionary<Document, SessionState>();

        [CommandMethod("PROCESSBOUNDARY_CC", CommandFlags.Session)]
        public void ProcessBoundaryCC()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            if (_sessions.ContainsKey(doc))
            {
                ed.WriteMessage("\n⚠ PROCESSBOUNDARY_CC is already running in this drawing.");
                return;
            }

            // ✅ Ask settings using WinForms dialog (reliable)
            var dlg = new BoundarySettingsForm(defaultSpacing: 100.0, defaultOrientation: BarsOrientation.Horizontal);
            WinForms.DialogResult dr = AcAp.ShowModalDialog(dlg);

            if (dr != WinForms.DialogResult.OK)
            {
                ed.WriteMessage("\n❌ Cancelled.");
                return;
            }

            double spacing = dlg.Spacing;
            BarsOrientation orientation = dlg.Orientation;

            if (spacing <= 0)
            {
                ed.WriteMessage("\n❌ Invalid spacing.");
                return;
            }

            var st = new SessionState
            {
                Doc = doc,
                Ed = ed,
                Db = db,
                Spacing = spacing,
                Orientation = orientation
            };
            _sessions[doc] = st;

            ed.WriteMessage($"\n✅ Spacing: {spacing}");
            ed.WriteMessage($"\n✅ Orientation: {orientation}");
            ed.WriteMessage("\nNow draw boundary using PLINE (type C to close OR end at start point), then press Enter...");

            StartPline(st);
        }

        // -------------------- PLINE CAPTURE --------------------
        private static void StartPline(SessionState st)
        {
            // Capture LAST polyline created during PLINE
            st.AppendedHandler = (sender, e) =>
            {
                if (e.DBObject is Entity ent && ent.OwnerId == st.Db.CurrentSpaceId)
                {
                    if (ent is Polyline || ent is Polyline2d || ent is Polyline3d)
                    {
                        st.CreatedBoundaryId = ent.ObjectId;
                    }
                }
            };
            st.Db.ObjectAppended += st.AppendedHandler;

            st.EndedHandler = (s, e) =>
            {
                if (!IsPlineCommand(e)) return;

                DetachCommandHandlers(st);

                // Run after PLINE is fully done
                st.IdleHandler = (ss, ee) =>
                {
                    AcAp.Idle -= st.IdleHandler;
                    st.IdleHandler = null;

                    RunBarsLogic(st);
                    Cleanup(st);
                };

                AcAp.Idle += st.IdleHandler;
            };

            st.CancelledHandler = (s, e) =>
            {
                if (!IsPlineCommand(e)) return;
                st.Ed.WriteMessage("\n❌ PLINE cancelled.");
                Cleanup(st);
            };

            st.FailedHandler = (s, e) =>
            {
                if (!IsPlineCommand(e)) return;
                st.Ed.WriteMessage("\n❌ PLINE failed.");
                Cleanup(st);
            };

            st.Doc.CommandEnded += st.EndedHandler;
            st.Doc.CommandCancelled += st.CancelledHandler;
            st.Doc.CommandFailed += st.FailedHandler;

            st.Doc.SendStringToExecute("_.PLINE ", true, false, false);
        }

        private static bool IsPlineCommand(CommandEventArgs e)
        {
            string cmd = (e.GlobalCommandName ?? "").Trim();
            return cmd.Equals("PLINE", StringComparison.OrdinalIgnoreCase) ||
                   cmd.Equals("_.PLINE", StringComparison.OrdinalIgnoreCase);
        }

        private static void DetachCommandHandlers(SessionState st)
        {
            if (st.EndedHandler != null) st.Doc.CommandEnded -= st.EndedHandler;
            if (st.CancelledHandler != null) st.Doc.CommandCancelled -= st.CancelledHandler;
            if (st.FailedHandler != null) st.Doc.CommandFailed -= st.FailedHandler;

            st.EndedHandler = null;
            st.CancelledHandler = null;
            st.FailedHandler = null;
        }

        private static void Cleanup(SessionState st)
        {
            if (st.IdleHandler != null)
            {
                AcAp.Idle -= st.IdleHandler;
                st.IdleHandler = null;
            }

            DetachCommandHandlers(st);

            if (st.AppendedHandler != null)
            {
                st.Db.ObjectAppended -= st.AppendedHandler;
                st.AppendedHandler = null;
            }

            _sessions.Remove(st.Doc);
        }

        // -------------------- MAIN BAR LOGIC --------------------
        private static void RunBarsLogic(SessionState st)
        {
            if (st.CreatedBoundaryId == ObjectId.Null)
            {
                st.Ed.WriteMessage("\n❌ No polyline created. Command cancelled.");
                return;
            }

            using (st.Doc.LockDocument())
            using (Transaction tr = st.Db.TransactionManager.StartTransaction())
            {
                Entity ent;
                try
                {
                    ent = (Entity)tr.GetObject(st.CreatedBoundaryId, OpenMode.ForWrite);
                }
                catch
                {
                    st.Ed.WriteMessage("\n❌ Created boundary not available.");
                    return;
                }

                Curve boundaryCurve = EnsureClosedCurve(st.Ed, ent);
                if (boundaryCurve == null) return;

                st.Ed.WriteMessage($"\n✅ Boundary Handle: {ent.Handle}");

                BlockTableRecord btr =
                    (BlockTableRecord)tr.GetObject(st.Db.CurrentSpaceId, OpenMode.ForWrite);

                Extents3d ext = boundaryCurve.GeometricExtents;

                double minX = ext.MinPoint.X;
                double maxX = ext.MaxPoint.X;
                double minY = ext.MinPoint.Y;
                double maxY = ext.MaxPoint.Y;

                double margin = Math.Max(1000.0, (maxX - minX) + (maxY - minY));
                double ptTol = 0.01;

                int barIndex = 1;
                double totalLength = 0.0;

                if (st.Orientation == BarsOrientation.Horizontal || st.Orientation == BarsOrientation.Both)
                {
                    GenerateHorizontalBars(st.Ed, tr, btr, boundaryCurve, minX, maxX, minY, maxY,
                        margin, st.Spacing, ptTol, ref barIndex, ref totalLength);
                }

                if (st.Orientation == BarsOrientation.Vertical || st.Orientation == BarsOrientation.Both)
                {
                    GenerateVerticalBars(st.Ed, tr, btr, boundaryCurve, minX, maxX, minY, maxY,
                        margin, st.Spacing, ptTol, ref barIndex, ref totalLength);
                }

                st.Ed.WriteMessage("\n====================");
                st.Ed.WriteMessage($"\nTotal Bars: {barIndex - 1}");
                st.Ed.WriteMessage($"\nTotal Length: {totalLength:F2} mm");
                st.Ed.WriteMessage("\n====================");

                tr.Commit();
            }
        }

        private static Curve EnsureClosedCurve(Editor ed, Entity ent)
        {
            if (ent is Polyline pl)
            {
                if (!pl.Closed)
                {
                    double closeTol = 0.5;
                    if (pl.StartPoint.DistanceTo(pl.EndPoint) <= closeTol)
                    {
                        pl.Closed = true;
                        ed.WriteMessage($"\n✅ Auto-closed polyline (tol={closeTol}).");
                    }
                }

                if (!pl.Closed)
                {
                    ed.WriteMessage("\n❌ Boundary must be closed (use C in PLINE or end at start point).");
                    return null;
                }

                return pl;
            }

            Curve c = ent as Curve;
            if (c == null || !c.Closed)
            {
                ed.WriteMessage("\n❌ Boundary must be a closed polyline/curve.");
                return null;
            }

            return c;
        }

        // -------------------- BAR GENERATION --------------------
        private static void GenerateHorizontalBars(
            Editor ed,
            Transaction tr,
            BlockTableRecord btr,
            Curve boundaryCurve,
            double minX,
            double maxX,
            double minY,
            double maxY,
            double margin,
            double spacing,
            double ptTol,
            ref int barIndex,
            ref double totalLength
        )
        {
            for (double y = minY; y <= maxY; y += spacing)
            {
                using (Line testLine = new Line(
                    new Point3d(minX - margin, y, 0),
                    new Point3d(maxX + margin, y, 0)
                ))
                {
                    Point3dCollection pts = new Point3dCollection();
                    testLine.IntersectWith(boundaryCurve, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);

                    if (pts.Count < 2) continue;

                    List<Point3d> intersections = new List<Point3d>();
                    foreach (Point3d p in pts) intersections.Add(p);

                    intersections.Sort((a, b) => a.X.CompareTo(b.X));
                    intersections = UniquePoints(intersections, ptTol);

                    int usable = intersections.Count - (intersections.Count % 2);
                    for (int i = 0; i < usable - 1; i += 2)
                    {
                        Line bar = new Line(intersections[i], intersections[i + 1]);
                        double len = bar.Length;
                        if (len <= 0.0001) { bar.Dispose(); continue; }

                        totalLength += len;
                        btr.AppendEntity(bar);
                        tr.AddNewlyCreatedDBObject(bar, true);

                        ed.WriteMessage($"\nBar {barIndex}: {len:F2} mm");
                        barIndex++;
                    }
                }
            }
        }

        private static void GenerateVerticalBars(
            Editor ed,
            Transaction tr,
            BlockTableRecord btr,
            Curve boundaryCurve,
            double minX,
            double maxX,
            double minY,
            double maxY,
            double margin,
            double spacing,
            double ptTol,
            ref int barIndex,
            ref double totalLength
        )
        {
            for (double x = minX; x <= maxX; x += spacing)
            {
                using (Line testLine = new Line(
                    new Point3d(x, minY - margin, 0),
                    new Point3d(x, maxY + margin, 0)
                ))
                {
                    Point3dCollection pts = new Point3dCollection();
                    testLine.IntersectWith(boundaryCurve, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);

                    if (pts.Count < 2) continue;

                    List<Point3d> intersections = new List<Point3d>();
                    foreach (Point3d p in pts) intersections.Add(p);

                    intersections.Sort((a, b) => a.Y.CompareTo(b.Y));
                    intersections = UniquePoints(intersections, ptTol);

                    int usable = intersections.Count - (intersections.Count % 2);
                    for (int i = 0; i < usable - 1; i += 2)
                    {
                        Line bar = new Line(intersections[i], intersections[i + 1]);
                        double len = bar.Length;
                        if (len <= 0.0001) { bar.Dispose(); continue; }

                        totalLength += len;
                        btr.AppendEntity(bar);
                        tr.AddNewlyCreatedDBObject(bar, true);

                        ed.WriteMessage($"\nBar {barIndex}: {len:F2} mm");
                        barIndex++;
                    }
                }
            }
        }

        private static List<Point3d> UniquePoints(List<Point3d> sortedPts, double tol)
        {
            List<Point3d> unique = new List<Point3d>();
            Point3d? last = null;

            foreach (var p in sortedPts)
            {
                if (last == null)
                {
                    unique.Add(p);
                    last = p;
                    continue;
                }

                if (((Point3d)last).DistanceTo(p) > tol)
                {
                    unique.Add(p);
                    last = p;
                }
            }

            return unique;
        }

        // -------------------- WINFORMS SETTINGS DIALOG --------------------
        private class BoundarySettingsForm : WinForms.Form
        {
            private readonly WinForms.NumericUpDown nudSpacing;
            private readonly WinForms.ComboBox cmbOrientation;

            public double Spacing => (double)nudSpacing.Value;

            public BarsOrientation Orientation
            {
                get
                {
                    switch (cmbOrientation.SelectedIndex)
                    {
                        case 1: return BarsOrientation.Vertical;
                        case 2: return BarsOrientation.Both;
                        default: return BarsOrientation.Horizontal;
                    }
                }
            }

            public BoundarySettingsForm(double defaultSpacing, BarsOrientation defaultOrientation)
            {
                Text = "Boundary Settings";
                FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                StartPosition = WinForms.FormStartPosition.CenterScreen;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                ClientSize = new Drawing.Size(360, 165);

                var lblSpacing = new WinForms.Label
                {
                    Text = "Spacing (mm):",
                    AutoSize = true,
                    Location = new Drawing.Point(20, 22)
                };

                nudSpacing = new WinForms.NumericUpDown
                {
                    Location = new Drawing.Point(140, 18),
                    Width = 180,
                    DecimalPlaces = 2,
                    Minimum = 0.01m,
                    Maximum = 1000000m,
                    Increment = 10m,
                    Value = (decimal)defaultSpacing
                };

                var lblOri = new WinForms.Label
                {
                    Text = "Orientation:",
                    AutoSize = true,
                    Location = new Drawing.Point(20, 62)
                };

                cmbOrientation = new WinForms.ComboBox
                {
                    Location = new Drawing.Point(140, 58),
                    Width = 180,
                    DropDownStyle = WinForms.ComboBoxStyle.DropDownList
                };
                cmbOrientation.Items.AddRange(new object[] { "Horizontal", "Vertical", "Both" });

                cmbOrientation.SelectedIndex =
                    defaultOrientation == BarsOrientation.Vertical ? 1 :
                    defaultOrientation == BarsOrientation.Both ? 2 : 0;

                var btnOk = new WinForms.Button
                {
                    Text = "OK",
                    DialogResult = WinForms.DialogResult.OK,
                    Location = new Drawing.Point(160, 110),
                    Width = 75
                };

                var btnCancel = new WinForms.Button
                {
                    Text = "Cancel",
                    DialogResult = WinForms.DialogResult.Cancel,
                    Location = new Drawing.Point(245, 110),
                    Width = 75
                };

                AcceptButton = btnOk;
                CancelButton = btnCancel;

                Controls.Add(lblSpacing);
                Controls.Add(nudSpacing);
                Controls.Add(lblOri);
                Controls.Add(cmbOrientation);
                Controls.Add(btnOk);
                Controls.Add(btnCancel);
            }
        }
    }
}
