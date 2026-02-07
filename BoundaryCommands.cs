using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

// WinForms / Drawing aliases to avoid AutoCAD name conflicts
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

using Newtonsoft.Json;

[assembly: CommandClass(typeof(CadBoundaryAutomation.BoundaryCommands))]

namespace CadBoundaryAutomation
{
    public class BoundaryCommands
    {
        private enum BarsOrientation { Horizontal, Vertical, Both }

        // ---------- JSON DTOs ----------
        private class PointJson
        {
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
        }

        private class BarJson
        {
            public int Index { get; set; }
            public string Orientation { get; set; } // "H" or "V"
            public double Length { get; set; }
            public string Handle { get; set; }
            public PointJson Start { get; set; }
            public PointJson End { get; set; }
        }

        private class BarsRunJson
        {
            public string DrawingName { get; set; }
            public string DrawingPath { get; set; }
            public DateTime CreatedAt { get; set; }

            public string BoundaryHandle { get; set; }

            public string Orientation { get; set; }
            public double SpacingH { get; set; }
            public double SpacingV { get; set; }

            public int TotalBars { get; set; }
            public double TotalLength { get; set; }

            public List<BarJson> Bars { get; set; }
        }

        private class SessionState
        {
            public Document Doc;
            public Editor Ed;
            public Database Db;

            public double SpacingH;
            public double SpacingV;
            public BarsOrientation Orientation;

            public ObjectId CreatedBoundaryId = ObjectId.Null;

            public ObjectEventHandler AppendedHandler;
            public CommandEventHandler EndedHandler;
            public CommandEventHandler CancelledHandler;
            public CommandEventHandler FailedHandler;
            public EventHandler IdleHandler;
        }

        private static readonly Dictionary<Document, SessionState> _sessions = new Dictionary<Document, SessionState>();

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

            // ✅ Dialog: if Both -> ask two spacings
            var dlg = new BoundarySettingsForm(
                defaultSpacingH: 100.0,
                defaultSpacingV: 100.0,
                defaultOrientation: BarsOrientation.Horizontal
            );

            WinForms.DialogResult dr = AcAp.ShowModalDialog(dlg);
            if (dr != WinForms.DialogResult.OK)
            {
                ed.WriteMessage("\n❌ Cancelled.");
                return;
            }

            BarsOrientation orientation = dlg.Orientation;
            double spacingH = dlg.SpacingH;
            double spacingV = dlg.SpacingV;

            // Validate
            if (orientation == BarsOrientation.Horizontal && spacingH <= 0)
            {
                ed.WriteMessage("\n❌ Invalid horizontal spacing.");
                return;
            }
            if (orientation == BarsOrientation.Vertical && spacingV <= 0)
            {
                ed.WriteMessage("\n❌ Invalid vertical spacing.");
                return;
            }
            if (orientation == BarsOrientation.Both && (spacingH <= 0 || spacingV <= 0))
            {
                ed.WriteMessage("\n❌ Invalid spacing(s) for Both.");
                return;
            }

            var st = new SessionState
            {
                Doc = doc,
                Ed = ed,
                Db = db,
                Orientation = orientation,
                SpacingH = spacingH,
                SpacingV = spacingV
            };
            _sessions[doc] = st;

            ed.WriteMessage($"\n✅ Orientation: {orientation}");
            if (orientation == BarsOrientation.Horizontal) ed.WriteMessage($"\n✅ Spacing(H): {spacingH}");
            else if (orientation == BarsOrientation.Vertical) ed.WriteMessage($"\n✅ Spacing(V): {spacingV}");
            else ed.WriteMessage($"\n✅ Spacing(H): {spacingH} | Spacing(V): {spacingV}");

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
                        st.CreatedBoundaryId = ent.ObjectId;
                }
            };
            st.Db.ObjectAppended += st.AppendedHandler;

            st.EndedHandler = (s, e) =>
            {
                if (!IsPlineCommand(e)) return;

                DetachCommandHandlers(st);

                // Run after PLINE fully ends
                st.IdleHandler = (ss, ee) =>
                {
                    AcAp.Idle -= st.IdleHandler;
                    st.IdleHandler = null;

                    RunBarsLogicAndSaveJson(st);
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

        // -------------------- MAIN BAR LOGIC + JSON SAVE --------------------
        private static void RunBarsLogicAndSaveJson(SessionState st)
        {
            if (st.CreatedBoundaryId == ObjectId.Null)
            {
                st.Ed.WriteMessage("\n❌ No polyline created. Command cancelled.");
                return;
            }

            var bars = new List<BarJson>();
            string boundaryHandle = "";
            int totalBars = 0;
            double totalLength = 0.0;

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

                boundaryHandle = ent.Handle.ToString();
                st.Ed.WriteMessage($"\n✅ Boundary Handle: {boundaryHandle}");

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

                if (st.Orientation == BarsOrientation.Horizontal || st.Orientation == BarsOrientation.Both)
                {
                    GenerateHorizontalBars(
                        st.Ed, tr, btr, boundaryCurve,
                        minX, maxX, minY, maxY, margin,
                        st.SpacingH, ptTol,
                        ref barIndex, ref totalLength,
                        bars
                    );
                }

                if (st.Orientation == BarsOrientation.Vertical || st.Orientation == BarsOrientation.Both)
                {
                    GenerateVerticalBars(
                        st.Ed, tr, btr, boundaryCurve,
                        minX, maxX, minY, maxY, margin,
                        st.SpacingV, ptTol,
                        ref barIndex, ref totalLength,
                        bars
                    );
                }

                totalBars = barIndex - 1;

                st.Ed.WriteMessage("\n====================");
                st.Ed.WriteMessage($"\nTotal Bars: {totalBars}");
                st.Ed.WriteMessage($"\nTotal Length: {totalLength:F2} mm");
                st.Ed.WriteMessage("\n====================");

                tr.Commit();
            }

            // ✅ Save JSON to file
            try
            {
                string jsonPath = GetDefaultJsonPath(st.Db);
                var run = new BarsRunJson
                {
                    DrawingName = Path.GetFileName(st.Db.Filename ?? st.Doc.Name),
                    DrawingPath = st.Db.Filename ?? "",
                    CreatedAt = DateTime.Now,

                    BoundaryHandle = boundaryHandle,

                    Orientation = st.Orientation.ToString(),
                    SpacingH = st.SpacingH,
                    SpacingV = st.SpacingV,

                    TotalBars = totalBars,
                    TotalLength = totalLength,

                    Bars = bars
                };

                string json = JsonConvert.SerializeObject(run, Formatting.Indented);
                File.WriteAllText(jsonPath, json);

                st.Ed.WriteMessage($"\n✅ Bars JSON saved: {jsonPath}");
            }
            catch (System.Exception ex)
            {
                st.Ed.WriteMessage($"\n❌ JSON save failed: {ex.Message}");
            }
        }

        private static string GetDefaultJsonPath(Database db)
        {
            string dwgPath = db.Filename;

            string folder;
            string name;

            if (!string.IsNullOrWhiteSpace(dwgPath))
            {
                folder = Path.GetDirectoryName(dwgPath);
                name = Path.GetFileNameWithoutExtension(dwgPath);
            }
            else
            {
                folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                name = "Drawing";
            }

            return Path.Combine(folder, name + "_bars.json");
        }

        private static Curve EnsureClosedCurve(Editor ed, Entity ent)
        {
            if (ent is Polyline pl)
            {
                if (!pl.Closed)
                {
                    double closeTol = 0.5; // drawing units (mm)
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

        // -------------------- BAR GENERATION (WITH JSON STORAGE) --------------------
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
            ref double totalLength,
            List<BarJson> store
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

                        store.Add(new BarJson
                        {
                            Index = barIndex,
                            Orientation = "H",
                            Length = len,
                            Handle = bar.Handle.ToString(),
                            Start = ToPointJson(bar.StartPoint),
                            End = ToPointJson(bar.EndPoint)
                        });

                        ed.WriteMessage($"\nH-Bar {barIndex}: {len:F2} mm");
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
            ref double totalLength,
            List<BarJson> store
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

                        store.Add(new BarJson
                        {
                            Index = barIndex,
                            Orientation = "V",
                            Length = len,
                            Handle = bar.Handle.ToString(),
                            Start = ToPointJson(bar.StartPoint),
                            End = ToPointJson(bar.EndPoint)
                        });

                        ed.WriteMessage($"\nV-Bar {barIndex}: {len:F2} mm");
                        barIndex++;
                    }
                }
            }
        }

        private static PointJson ToPointJson(Point3d p)
        {
            return new PointJson { X = p.X, Y = p.Y, Z = p.Z };
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
            private readonly WinForms.NumericUpDown nudSpacingH;
            private readonly WinForms.NumericUpDown nudSpacingV;
            private readonly WinForms.ComboBox cmbOrientation;

            public double SpacingH => (double)nudSpacingH.Value;
            public double SpacingV => (double)nudSpacingV.Value;

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

            public BoundarySettingsForm(double defaultSpacingH, double defaultSpacingV, BarsOrientation defaultOrientation)
            {
                Text = "Boundary Settings";
                FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                StartPosition = WinForms.FormStartPosition.CenterScreen;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                ClientSize = new Drawing.Size(420, 210);

                var lblOri = new WinForms.Label
                {
                    Text = "Orientation:",
                    AutoSize = true,
                    Location = new Drawing.Point(20, 20)
                };

                cmbOrientation = new WinForms.ComboBox
                {
                    Location = new Drawing.Point(200, 16),
                    Width = 180,
                    DropDownStyle = WinForms.ComboBoxStyle.DropDownList
                };
                cmbOrientation.Items.AddRange(new object[] { "Horizontal", "Vertical", "Both" });
                cmbOrientation.SelectedIndex =
                    defaultOrientation == BarsOrientation.Vertical ? 1 :
                    defaultOrientation == BarsOrientation.Both ? 2 : 0;

                cmbOrientation.SelectedIndexChanged += (s, e) => UpdateEnableState();

                var lblH = new WinForms.Label
                {
                    Text = "Horizontal spacing (mm):",
                    AutoSize = true,
                    Location = new Drawing.Point(20, 65)
                };

                nudSpacingH = new WinForms.NumericUpDown
                {
                    Location = new Drawing.Point(200, 61),
                    Width = 180,
                    DecimalPlaces = 2,
                    Minimum = 0.01m,
                    Maximum = 1000000m,
                    Increment = 10m,
                    Value = (decimal)Math.Max(0.01, defaultSpacingH)
                };

                var lblV = new WinForms.Label
                {
                    Text = "Vertical spacing (mm):",
                    AutoSize = true,
                    Location = new Drawing.Point(20, 105)
                };

                nudSpacingV = new WinForms.NumericUpDown
                {
                    Location = new Drawing.Point(200, 101),
                    Width = 180,
                    DecimalPlaces = 2,
                    Minimum = 0.01m,
                    Maximum = 1000000m,
                    Increment = 10m,
                    Value = (decimal)Math.Max(0.01, defaultSpacingV)
                };

                var btnOk = new WinForms.Button
                {
                    Text = "OK",
                    DialogResult = WinForms.DialogResult.OK,
                    Location = new Drawing.Point(225, 150),
                    Width = 75
                };

                var btnCancel = new WinForms.Button
                {
                    Text = "Cancel",
                    DialogResult = WinForms.DialogResult.Cancel,
                    Location = new Drawing.Point(305, 150),
                    Width = 75
                };

                AcceptButton = btnOk;
                CancelButton = btnCancel;

                Controls.Add(lblOri);
                Controls.Add(cmbOrientation);
                Controls.Add(lblH);
                Controls.Add(nudSpacingH);
                Controls.Add(lblV);
                Controls.Add(nudSpacingV);
                Controls.Add(btnOk);
                Controls.Add(btnCancel);

                UpdateEnableState();
            }

            private void UpdateEnableState()
            {
                var ori = Orientation;

                if (ori == BarsOrientation.Horizontal)
                {
                    nudSpacingH.Enabled = true;
                    nudSpacingV.Enabled = false;
                }
                else if (ori == BarsOrientation.Vertical)
                {
                    nudSpacingH.Enabled = false;
                    nudSpacingV.Enabled = true;
                }
                else
                {
                    nudSpacingH.Enabled = true;
                    nudSpacingV.Enabled = true;
                }
            }
        }
    }
}
