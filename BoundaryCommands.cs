using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(AutoCadBoundaryProcessor.BoundaryCommands))]

namespace AutoCadBoundaryProcessor
{
    public class BoundaryCommands
    {
        private const double SPACING = 100.0; // C/C spacing in mm
        private const double CLOSE_TOLERANCE = 0.5; // Auto-close tolerance in mm
        private const double MARGIN_MULTIPLIER = 1.5; // Safety margin for intersection tests

        [CommandMethod("PROCESSBOUNDARY_CC")]
        public void ProcessBoundaryCC()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\nDraw a CLOSED boundary using PLINE (type C to close, or end at start point), then press Enter...");

            ObjectId createdBoundaryId = GetBoundaryFromPline(db, ed);

            if (createdBoundaryId == ObjectId.Null)
            {
                ed.WriteMessage("\n❌ No polyline created. Command cancelled.");
                return;
            }

            // Process the boundary and generate bars
            ProcessBoundaryAndGenerateBars(db, ed, createdBoundaryId);
        }

        /// <summary>
        /// Captures the polyline created by the interactive PLINE command
        /// </summary>
        private ObjectId GetBoundaryFromPline(Database db, Editor ed)
        {
            ObjectId createdBoundaryId = ObjectId.Null;
            bool foundPolyline = false;

            // Event handler to capture the first polyline added to current space
            ObjectEventHandler appendedHandler = (sender, e) =>
            {
                if (foundPolyline)
                    return;

                if (e.DBObject is Entity ent && ent.OwnerId == db.CurrentSpaceId)
                {
                    // Check for all polyline types
                    if (ent is Polyline || ent is Polyline2d || ent is Polyline3d)
                    {
                        createdBoundaryId = ent.ObjectId;
                        foundPolyline = true;
                    }
                }
            };

            db.ObjectAppended += appendedHandler;

            try
            {
                // Execute PLINE command interactively
                ed.Command("_.PLINE");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                // User cancelled with ESC
                if (ex.ErrorStatus == ErrorStatus.UserBreak)
                {
                    ed.WriteMessage("\n⚠️ Command cancelled by user.");
                    return ObjectId.Null;
                }

                ed.WriteMessage($"\n❌ Error while running PLINE: {ex.Message}");
                return ObjectId.Null;
            }
            finally
            {
                // Always unsubscribe from event
                db.ObjectAppended -= appendedHandler;
            }

            return createdBoundaryId;
        }

        /// <summary>
        /// Validates and auto-closes the boundary curve
        /// </summary>
        private Curve ValidateAndPrepareBoundary(Transaction tr, ObjectId boundaryId, Editor ed)
        {
            Entity ent = tr.GetObject(boundaryId, OpenMode.ForWrite) as Entity;
            if (ent == null)
            {
                ed.WriteMessage("\n❌ Invalid entity reference.");
                return null;
            }

            // Handle lightweight polyline (most common)
            if (ent is Polyline pl)
            {
                if (!pl.Closed)
                {
                    // Check if start and end points are close enough to auto-close
                    double distance = pl.StartPoint.DistanceTo(pl.EndPoint);

                    if (distance <= CLOSE_TOLERANCE)
                    {
                        pl.Closed = true;
                        ed.WriteMessage($"\n✅ Auto-closed polyline (gap: {distance:F3} mm, tolerance: {CLOSE_TOLERANCE} mm).");
                    }
                    else
                    {
                        ed.WriteMessage($"\n❌ Boundary must be closed. Current gap: {distance:F3} mm (use 'C' to close or end at start point).");
                        return null;
                    }
                }

                return pl;
            }

            // Handle other curve types
            Curve boundaryCurve = ent as Curve;

            if (boundaryCurve == null)
            {
                ed.WriteMessage("\n❌ Entity must be a curve (polyline, spline, circle, etc.).");
                return null;
            }

            if (!boundaryCurve.Closed)
            {
                ed.WriteMessage("\n❌ Boundary must be a closed curve.");
                return null;
            }

            return boundaryCurve;
        }

        /// <summary>
        /// Main processing logic: validates boundary and generates intersection bars
        /// </summary>
        private void ProcessBoundaryAndGenerateBars(Database db, Editor ed, ObjectId boundaryId)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Validate and prepare the boundary
                    Curve boundaryCurve = ValidateAndPrepareBoundary(tr, boundaryId, ed);

                    if (boundaryCurve == null)
                    {
                        tr.Abort();
                        return;
                    }

                    ed.WriteMessage($"\n✅ Boundary Handle: {boundaryCurve.Handle}");

                    // Open current space for writing
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    if (btr == null)
                    {
                        ed.WriteMessage("\n❌ Could not access current space.");
                        tr.Abort();
                        return;
                    }

                    // Generate bars using scanline algorithm
                    BoundaryStats stats = GenerateBarsInBoundary(boundaryCurve, btr, tr, ed);

                    // Display summary
                    DisplaySummary(ed, stats);

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n❌ Error during processing: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        /// <summary>
        /// Generates horizontal bars inside the boundary using scanline algorithm
        /// </summary>
        private BoundaryStats GenerateBarsInBoundary(
            Curve boundaryCurve,
            BlockTableRecord btr,
            Transaction tr,
            Editor ed)
        {
            BoundaryStats stats = new BoundaryStats();

            // Get boundary extents
            Extents3d ext = boundaryCurve.GeometricExtents;
            double minX = ext.MinPoint.X;
            double maxX = ext.MaxPoint.X;
            double minY = ext.MinPoint.Y;
            double maxY = ext.MaxPoint.Y;

            // Calculate margin for test lines (extend beyond boundary)
            double boundaryWidth = maxX - minX;
            double boundaryHeight = maxY - minY;
            double margin = Math.Max(boundaryWidth, boundaryHeight) * MARGIN_MULTIPLIER;

            ed.WriteMessage($"\n📐 Boundary extents: Width={boundaryWidth:F2} mm, Height={boundaryHeight:F2} mm");
            ed.WriteMessage($"\n📏 Generating bars with {SPACING:F2} mm spacing...\n");

            // Scanline algorithm: horizontal sweep from bottom to top
            for (double y = minY; y <= maxY; y += SPACING)
            {
                List<Line> barsAtY = GenerateBarsAtScanline(
                    boundaryCurve,
                    y,
                    minX,
                    maxX,
                    margin,
                    ref stats);

                // Add bars to drawing
                foreach (Line bar in barsAtY)
                {
                    btr.AppendEntity(bar);
                    tr.AddNewlyCreatedDBObject(bar, true);

                    ed.WriteMessage($"Bar {stats.BarCount}: {bar.Length:F2} mm");
                }
            }

            return stats;
        }

        /// <summary>
        /// Generates bars at a specific Y-coordinate scanline
        /// </summary>
        private List<Line> GenerateBarsAtScanline(
            Curve boundaryCurve,
            double y,
            double minX,
            double maxX,
            double margin,
            ref BoundaryStats stats)
        {
            List<Line> bars = new List<Line>();

            // Create horizontal test line across the entire width
            using (Line testLine = new Line(
                new Point3d(minX - margin, y, 0),
                new Point3d(maxX + margin, y, 0)))
            {
                Point3dCollection intersectionPoints = new Point3dCollection();

                // Find all intersection points with boundary
                boundaryCurve.IntersectWith(
                    testLine,
                    Intersect.OnBothOperands,
                    intersectionPoints,
                    IntPtr.Zero,
                    IntPtr.Zero);

                // Need at least 2 points to create a bar
                if (intersectionPoints.Count < 2)
                {
                    return bars;
                }

                // Convert to sorted list (left to right)
                List<Point3d> sortedPoints = new List<Point3d>();
                foreach (Point3d p in intersectionPoints)
                {
                    sortedPoints.Add(p);
                }
                sortedPoints.Sort((a, b) => a.X.CompareTo(b.X));

                // Create bars by pairing consecutive points (0-1, 2-3, 4-5, ...)
                for (int i = 0; i < sortedPoints.Count - 1; i += 2)
                {
                    Line bar = new Line(sortedPoints[i], sortedPoints[i + 1]);
                    double length = bar.Length;

                    // Filter out degenerate bars (too short)
                    if (length > 0.001) // 0.001 mm threshold
                    {
                        bars.Add(bar);
                        stats.BarCount++;
                        stats.TotalLength += length;
                    }
                    else
                    {
                        bar.Dispose();
                    }
                }
            }

            return bars;
        }

        /// <summary>
        /// Displays the summary statistics
        /// </summary>
        private void DisplaySummary(Editor ed, BoundaryStats stats)
        {
            ed.WriteMessage("\n====================");
            ed.WriteMessage($"\n✅ Total Bars Created: {stats.BarCount}");
            ed.WriteMessage($"\n📏 Total Length: {stats.TotalLength:F2} mm ({stats.TotalLength / 1000:F3} m)");
            ed.WriteMessage($"\n📊 Average Bar Length: {(stats.BarCount > 0 ? stats.TotalLength / stats.BarCount : 0):F2} mm");
            ed.WriteMessage("\n====================");
        }

        /// <summary>
        /// Helper class to track boundary processing statistics
        /// </summary>
        private class BoundaryStats
        {
            public int BarCount { get; set; } = 0;
            public double TotalLength { get; set; } = 0.0;
        }
    }
}
