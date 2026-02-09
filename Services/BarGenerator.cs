using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using CadBoundaryAutomation.Models;

namespace CadBoundaryAutomation.Services
{
    public static class BarGenerator
    {
        public static void GenerateHorizontalBars(
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
                            Orientation = "Horizontal",
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

        public static void GenerateVerticalBars(
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
                            Orientation = "Vertical",
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

        public static List<Point3d> UniquePoints(List<Point3d> sortedPts, double tol)
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

        public static PointJson ToPointJson(Point3d p)
        {
            return new PointJson { X = p.X, Y = p.Y, Z = p.Z };
        }
    }
}
