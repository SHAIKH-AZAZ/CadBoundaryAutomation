using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(AutoCadBoundaryProcessor.BoundaryCommands))]
// this is created by azaz
namespace AutoCadBoundaryProcessor
{
    public class BoundaryCommands
    {
        [CommandMethod("PROCESSBOUNDARY_CC")]
        public void ProcessBoundaryCC()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage(
                "\nDraw a CLOSED boundary using PLINE (type C to close, or end at start point), then press Enter..."
            );

            ObjectId createdBoundaryId = ObjectId.Null;

            // Capture the first polyline created during the PLINE command
            ObjectEventHandler appendedHandler = (sender, e) =>
            {
                if (createdBoundaryId != ObjectId.Null) return;

                if (e.DBObject is Entity ent && ent.OwnerId == db.CurrentSpaceId)
                {
                    if (ent is Polyline || ent is Polyline2d || ent is Polyline3d)
                    {
                        createdBoundaryId = ent.ObjectId;
                    }
                }
            };

            db.ObjectAppended += appendedHandler;

            try
            {
                // Run PLINE interactively (returns when user finishes)
                ed.Command("_.PLINE");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                if (ex.ErrorStatus == ErrorStatus.UserBreak)
                    return;

                ed.WriteMessage($"\n❌ Error while running PLINE: {ex.Message}");
                return;
            }
            finally
            {
                db.ObjectAppended -= appendedHandler;
            }

            if (createdBoundaryId == ObjectId.Null)
            {
                ed.WriteMessage("\n❌ No polyline created. Command cancelled.");
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open for write because we may set Closed=true
                Entity ent = (Entity)tr.GetObject(createdBoundaryId, OpenMode.ForWrite);

                // Ensure we end up with a closed Curve
                Curve boundaryCurve = null;

                // Auto-close only for LWPolyline (most common PLINE result)
                if (ent is Polyline pl)
                {
                    if (!pl.Closed)
                    {
                        // Tolerance in drawing units (mm if drawing is mm)
                        double closeTol = 0.5;

                        if (pl.StartPoint.DistanceTo(pl.EndPoint) <= closeTol)
                        {
                            pl.Closed = true; // auto-close
                            ed.WriteMessage($"\n✅ Auto-closed polyline (tol={closeTol}).");
                        }
                    }

                    if (!pl.Closed)
                    {
                        ed.WriteMessage(
                            "\n❌ Boundary must be closed (use C in PLINE or end at start point)."
                        );
                        return;
                    }

                    boundaryCurve = pl;
                }
                else
                {
                    // For other polyline types / curves
                    boundaryCurve = ent as Curve;

                    if (boundaryCurve == null || !boundaryCurve.Closed)
                    {
                        ed.WriteMessage("\n❌ Boundary must be a closed polyline/curve.");
                        return;
                    }
                }

                ed.WriteMessage($"\n✅ Boundary Handle: {ent.Handle}");

                BlockTableRecord btr =
                    (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Get boundary extents
                Extents3d ext = boundaryCurve.GeometricExtents;

                double minX = ext.MinPoint.X;
                double maxX = ext.MaxPoint.X;
                double minY = ext.MinPoint.Y;
                double maxY = ext.MaxPoint.Y;

                const double spacing = 100.0; // 100 mm C/C

                int barIndex = 1;
                double totalLength = 0.0;

                // Add extra margin so test line surely crosses boundary
                double margin = Math.Max(1000.0, (maxX - minX) + (maxY - minY));

                // Scanline fill (horizontal bars)
                for (double y = minY; y <= maxY; y += spacing)
                {
                    using (Line testLine = new Line(
                        new Point3d(minX - margin, y, 0),
                        new Point3d(maxX + margin, y, 0)
                    ))
                    {
                        Point3dCollection pts = new Point3dCollection();

                        testLine.IntersectWith(
                            boundaryCurve,
                            Intersect.OnBothOperands,
                            pts,
                            IntPtr.Zero,
                            IntPtr.Zero
                        );

                        if (pts.Count < 2)
                            continue;

                        List<Point3d> intersections = new List<Point3d>();
                        foreach (Point3d p in pts)
                            intersections.Add(p);

                        // Sort intersections from left to right
                        intersections.Sort((a, b) => a.X.CompareTo(b.X));

                        // Pair points (0-1, 2-3, ...)
                        for (int i = 0; i < intersections.Count - 1; i += 2)
                        {
                            Line bar = new Line(intersections[i], intersections[i + 1]);

                            double len = bar.Length;
                            if (len <= 0.0001)
                            {
                                bar.Dispose();
                                continue;
                            }

                            totalLength += len;

                            btr.AppendEntity(bar);
                            tr.AddNewlyCreatedDBObject(bar, true);

                            ed.WriteMessage($"\nBar {barIndex}: {len:F2} mm");
                            barIndex++;
                        }
                    }
                }

                ed.WriteMessage("\n====================");
                ed.WriteMessage($"\nTotal Bars: {barIndex - 1}");
                ed.WriteMessage($"\nTotal Length: {totalLength:F2} mm");
                ed.WriteMessage("\n====================");

                tr.Commit();
            }
        }
    }
}
