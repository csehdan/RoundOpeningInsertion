using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RoundOpeningInsertion
{
	public static class Helpers
	{
		public static IEnumerable<Face> FindWallFace(Wall wall)
        {
            var normalFaces = new List<Face>();

            var opt = new Options();
            opt.ComputeReferences = true;
            opt.DetailLevel = ViewDetailLevel.Fine;

            var e = wall.get_Geometry(opt);

            foreach (GeometryObject obj in e)
            {
                var solid = obj as Solid;

                if (solid != null && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        var pf = face as PlanarFace;

                        if (pf == null)
                        {
                            continue;
                        }

                        if ((int)pf.FaceNormal.Z == 0
                            && (Math.Abs(wall.Orientation.X) == Math.Abs(pf.FaceNormal.X) || Math.Abs(wall.Orientation.Y) == Math.Abs(pf.FaceNormal.Y)))
                        {
                            normalFaces.Add(pf);
                        }
                    }
                }
            }
            return normalFaces;
        }

        public static Family FindOrLoadFamily(Document doc, string familyName)
        {
            var a = new FilteredElementCollector(doc).OfClass(typeof(Family));
            var family = a.FirstOrDefault(e => e.Name.Equals(familyName)) as Family;

            if (family == null)
            {
                var workingDirectory = Environment.CurrentDirectory;
                var projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
                var FamilyPath = Path.Combine(projectDirectory, familyName + ".rfa");

                if (!File.Exists(FamilyPath))
                {
                    return null;
                }

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Load Family");
                    doc.LoadFamily(FamilyPath, out family);
                    tx.Commit();
                }
            }

            if (family == null)
            {
                throw new Exception("Failed to load family");
            }

            return family;
        }

        public static FamilySymbol ResolveFamilySymbol(Family family, Document doc)
        {
            var familySymbol = GetFamilySymbol(family);

            if (familySymbol == null)
            {
                throw new Exception("Family symbol not found"); ;
            }

            if (!familySymbol.IsActive)
            {
                using (var tx = new Transaction(doc))
                {
                    tx.Start("Activate family symbol");
                    familySymbol.Activate();
                    tx.Commit();
                }
            }

            return familySymbol;
        }

        public static FamilySymbol GetFamilySymbol(Family family)
        {
            var elementId = family.GetFamilySymbolIds().FirstOrDefault();

            if (elementId == null)
            {
                return null;
            }

            var familySymbol = family.Document.GetElement(elementId) as FamilySymbol;

            return familySymbol;
        }

        public static Curve FindDuctCurve(Duct duct)
        {
            var list = new List<XYZ>();

            var csi = duct.ConnectorManager.Connectors.ForwardIterator();
            while (csi.MoveNext())
            {
                var conn = csi.Current as Connector;
                list.Add(conn.Origin);
            }
            var curve = Line.CreateBound(list.ElementAt(0), list.ElementAt(1)) as Curve;
            curve.MakeUnbound();

            return curve;
        }

        public static double GetDiameter(double horizontalDiff, double verticalDiff, double width, double diameter)
        {
            var diff = Math.Sqrt(verticalDiff * verticalDiff + horizontalDiff * horizontalDiff);
            if (diff > 0)
            {
                var coff = diff / width;
                var edge = diameter * coff;
                var newDiameter = Math.Sqrt(diameter * diameter + edge * edge);
                return newDiameter + diff;
            }
            return diameter;
        }

        public static XYZ GetRefDir(Face face)
        {
            var bboxUV = face.GetBoundingBox();
            var center = (bboxUV.Max + bboxUV.Min) / 2d;
            var normal = face.ComputeNormal(center);
            return normal.CrossProduct(XYZ.BasisZ);
        }

        public static XYZ FindIntersection(Curve ductCurve, Face face)
        {
            var results = face.Intersect(ductCurve, out var intersectionR);

            XYZ intersectionResult = null;

            if (SetComparisonResult.Disjoint != results)
            {
                if (intersectionR != null)
                {
                    if (!intersectionR.IsEmpty)
                    {
                        intersectionResult = intersectionR.get_Item(0).XYZPoint;
                    }
                }
            }
            return intersectionResult;
        }
    }
}
