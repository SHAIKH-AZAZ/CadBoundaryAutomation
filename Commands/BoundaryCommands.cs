using System;
using System.Collections.Generic;
using System.IO;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

using CadBoundaryAutomation.Models;
using CadBoundaryAutomation.Services;
using CadBoundaryAutomation.UI;

[assembly: CommandClass(typeof(CadBoundaryAutomation.Commands.BoundaryCommands))]

namespace CadBoundaryAutomation.Commands
{
    public class BoundaryCommands
    {
        [CommandMethod("PROCESSBOUNDARY_CC", CommandFlags.Session)]
        public void ProcessBoundaryCC()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // 1) Dialog first (reliable)
            var dlg = new BoundarySettingsForm(100.0, 100.0, BarsOrientation.Vertical);
            var dr = AcAp.ShowModalDialog(dlg);

            if (dr != System.Windows.Forms.DialogResult.OK)
            {
                ed.WriteMessage("\n❌ Cancelled.");
                return;
            }

            BarsOrientation orientation = dlg.Orientation;

            double spacingH = dlg.SpacingH;
            double spacingV = dlg.SpacingV;

            if (orientation == BarsOrientation.Horizontal && spacingH <= 0)
            {
                ed.WriteMessage("\n❌ Invalid H spacing.");
                return;
            }

            if (orientation == BarsOrientation.Vertical && spacingV <= 0)
            {
                ed.WriteMessage("\n❌ Invalid V spacing.");
                return;
            }

            if (orientation == BarsOrientation.Both && (spacingH <= 0 || spacingV <= 0))
            {
                ed.WriteMessage("\n❌ Invalid spacing(s).");
                return;
            }

            ed.WriteMessage($"\n✅ Orientation: {orientation}");
            ed.WriteMessage($"\n✅ H spacing: {spacingH} | V spacing: {spacingV}");
            ed.WriteMessage("\nNow draw boundary using PLINE (close it), then press Enter...");

            // 2) Run PLINE, then continue when done
            var runner = new PlineRunner(doc);

            runner.Cancelled += () =>
            {
                ed.WriteMessage("\n❌ PLINE cancelled/failed.");
            };

            runner.Completed += () =>
            {
                if (runner.CreatedBoundaryId == ObjectId.Null)
                {
                    ed.WriteMessage("\n❌ No polyline created.");
                    return;
                }

                RunBarsAndSaveJson(doc, runner.CreatedBoundaryId, orientation, spacingH, spacingV);
            };

            runner.Start();
        }

        private void RunBarsAndSaveJson(
            Document doc,
            ObjectId boundaryId,
            BarsOrientation orientation,
            double spacingH,
            double spacingV)
        {
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Grouped by: Orientation + rounded length (2 decimals)
            var groupedBars = new Dictionary<string, BarJson>();

            double totalLength = 0.0;
            int barIndex = 1;
            string boundaryHandle = "";

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var ent = (Entity)tr.GetObject(boundaryId, OpenMode.ForWrite);
                var curve = EnsureClosedCurve(ed, ent);
                if (curve == null) return;

                boundaryHandle = ent.Handle.ToString();
                ed.WriteMessage($"\n✅ Boundary Handle: {boundaryHandle}");

                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                Extents3d ext = curve.GeometricExtents;
                double minX = ext.MinPoint.X;
                double maxX = ext.MaxPoint.X;
                double minY = ext.MinPoint.Y;
                double maxY = ext.MaxPoint.Y;

                double margin = Math.Max(1000.0, (maxX - minX) + (maxY - minY));
                double ptTol = 0.01;

                if (orientation == BarsOrientation.Horizontal || orientation == BarsOrientation.Both)
                {
                    BarGenerator.GenerateHorizontalBars(
                        ed, tr, btr, curve,
                        minX, maxX, minY, maxY,
                        margin, spacingH, ptTol,
                        ref barIndex, ref totalLength,
                        groupedBars
                    );
                }

                if (orientation == BarsOrientation.Vertical || orientation == BarsOrientation.Both)
                {
                    BarGenerator.GenerateVerticalBars(
                        ed, tr, btr, curve,
                        minX, maxX, minY, maxY,
                        margin, spacingV, ptTol,
                        ref barIndex, ref totalLength,
                        groupedBars
                    );
                }

                tr.Commit();
            }

            // Convert groups -> list for JSON
            var barsList = new List<BarJson>(groupedBars.Values);

            // Optional: sort output (nice JSON)
            barsList.Sort((a, b) =>
            {
                int o = string.Compare(a.Orientation, b.Orientation, StringComparison.OrdinalIgnoreCase);
                if (o != 0) return o;
                return a.Length.CompareTo(b.Length);
            });

            // Total bars = sum of repetition counts
            int totalBars = 0;
            foreach (var item in barsList)
                totalBars += item.Repetition;

            // Round total length to 2 decimals
            totalLength = Math.Round(totalLength, 2, MidpointRounding.AwayFromZero);

            if (totalBars == 0)
                ed.WriteMessage("\n⚠ No bars were generated inside the boundary (check spacing / boundary).");

            var run = new BarsRunJson
            {
                DrawingName = Path.GetFileName(db.Filename ?? doc.Name),
                DrawingPath = db.Filename ?? "",
                CreatedAt = DateTime.Now,
                BoundaryHandle = boundaryHandle,
                Orientation = orientation.ToString(),
                SpacingH = spacingH,
                SpacingV = spacingV,
                TotalBars = totalBars,
                TotalLength = totalLength,
                Bars = barsList
            };

            try
            {
                string jsonPath = JsonExporter.GetDefaultJsonPath(db);
                JsonExporter.Save(jsonPath, run);
                ed.WriteMessage($"\n✅ Bars JSON saved: {jsonPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ JSON save failed: {ex.Message}");
            }

            ed.WriteMessage("\n====================");
            ed.WriteMessage($"\nTotal Bars: {totalBars}");
            ed.WriteMessage($"\nTotal Length: {totalLength:F2} mm");
            ed.WriteMessage("\n====================");
        }

        private static Curve EnsureClosedCurve(Editor ed, Entity ent)
        {
            if (ent is Polyline pl)
            {
                if (!pl.Closed)
                {
                    double closeTol = 0.5; // drawing units (mm)
                    if (pl.StartPoint.DistanceTo(pl.EndPoint) <= closeTol)
                        pl.Closed = true;
                }

                if (!pl.Closed)
                {
                    ed.WriteMessage("\n❌ Boundary must be closed.");
                    return null;
                }

                return pl;
            }

            var c = ent as Curve;
            if (c == null || !c.Closed)
            {
                ed.WriteMessage("\n❌ Boundary must be closed.");
                return null;
            }

            return c;
        }
    }
}
