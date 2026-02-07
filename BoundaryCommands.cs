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
        [CommandMethod("PROCESSBOUNDARY_CC")]
        public void ProcessBoundaryCC()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptPointResult ppr = ed.GetPoint("\nPick a point inside the shape: ");
            if (ppr.Status != PromptStatus.OK) return;
            Point3d anchor = ppr.Value;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBObjectCollection boundaryCurves = ed.TraceBoundary(anchor, false);
                if (boundaryCurves.Count == 0)
                {
                    ed.WriteMessage("\n❌ No closed boundary found.");
                    return;
                }

                BlockTableRecord btr =
                    (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Compute extents
                Extents3d ext = new Extents3d();
                bool first = true;

                foreach (DBObject obj in boundaryCurves)
                {
                    if (obj is Entity ent)
                    {
                        if (first)
                        {
                            ext = ent.GeometricExtents;
                            first = false;
                        }
                        else
                        {
                            ext.AddExtents(ent.GeometricExtents);
                        }
                    }
                }


                double minX = ext.MinPoint.X;
                double maxX = ext.MaxPoint.X;
                double minY = ext.MinPoint.Y;
                double maxY = ext.MaxPoint.Y;

                const double spacing = 100.0; // 100 mm C/C

                int barIndex = 1;
                double totalLength = 0;

                // 🔹 Generate horizontal lines
                for (double y = minY; y <= maxY; y += spacing)
                {
                    Line testLine = new Line(
                        new Point3d(minX - 1000, y, 0),
                        new Point3d(maxX + 1000, y, 0)
                    );

                    List<Point3d> intersections = new List<Point3d>();

                    foreach (Entity boundary in boundaryCurves)
                    {
                        if (boundary is Curve bc)
                        {
                            Point3dCollection pts = new Point3dCollection();
                            testLine.IntersectWith(
                                bc,
                                Intersect.OnBothOperands,
                                pts,
                                IntPtr.Zero,
                                IntPtr.Zero
                            );

                            foreach (Point3d p in pts)
                                intersections.Add(p);
                        }
                    }

                    if (intersections.Count < 2) continue;

                    intersections.Sort((a, b) => a.X.CompareTo(b.X));

                    for (int i = 0; i < intersections.Count - 1; i += 2)
                    {
                        Line bar = new Line(intersections[i], intersections[i + 1]);
                        double len = bar.Length;

                        totalLength += len;
                        ed.WriteMessage($"\nBar {barIndex}: {len:F2} mm");

                        btr.AppendEntity(bar);
                        tr.AddNewlyCreatedDBObject(bar, true);

                        barIndex++;
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