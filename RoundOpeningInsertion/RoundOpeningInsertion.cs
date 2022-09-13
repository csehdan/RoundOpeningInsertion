using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RoundOpeningInsertion
{
	[Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    public class RoundOpeningInsertion : IExternalCommand
    {
        private readonly string familyName = "M_Round Face Opening";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var doc = uiapp.ActiveUIDocument.Document;
            var family = FindOrLoadFamily(doc);

            if(family == null)
            {
                return Result.Failed;
            }
            var familySymbol = GetFamilySymbol(family);

            if(familySymbol == null)
            {
                return Result.Failed;
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

            var walls = new FilteredElementCollector(doc).OfClass(typeof(Wall));
            using (var t = new Transaction(doc, "Duct Wall Intersection"))
            {
                t.Start();

                foreach (Wall wall in walls)
                {
                    var bounding = wall.get_BoundingBox(null);
                    var outline = new Outline(bounding.Min, bounding.Max);
                    var bbfilter = new BoundingBoxIntersectsFilter(outline);
                    var ducts = new FilteredElementCollector(doc).OfClass(typeof(Duct));
                    var intersectedDucts = ducts.WherePasses(bbfilter).Cast<Duct>().ToList();

                    if (intersectedDucts.Count > 0)
                    {
                        var wallFaces = FindWallFace(wall);

                        if (wallFaces.Count == 2)
                        {
                            foreach (var duct in intersectedDucts)
                            {
                                var ductCurve = FindDuctCurve(duct);
                                var face = wallFaces[0];
                                var frontIntersection = FindIntersection(ductCurve, face);
                                var backIntersection = FindIntersection(ductCurve, wallFaces[1]);

                                if (frontIntersection is null || backIntersection is null)
                                {
                                    continue;
                                }

                                var frontRefDir = GetRefDir(face);
                                XYZ intersection = null;
                                var verticalDiff = frontIntersection.Z - backIntersection.Z;
                                var horizontalDiff = 0d;

                                if(Math.Abs(frontRefDir.Y) > Math.Abs(frontRefDir.X))
                                {
                                    horizontalDiff = frontIntersection.Y - backIntersection.Y;
                                    intersection = new XYZ(frontIntersection.X, frontIntersection.Y - (horizontalDiff / 2), frontIntersection.Z - (verticalDiff / 2));
                                }
                                else if(Math.Abs(frontRefDir.X) > Math.Abs(frontRefDir.Y))
                                {
                                    horizontalDiff = frontIntersection.X - backIntersection.X;
                                    intersection = new XYZ(frontIntersection.X - (horizontalDiff / 2), frontIntersection.Y , frontIntersection.Z - (verticalDiff / 2));
                                }

                                var instance = doc.Create.NewFamilyInstance(face, intersection, frontRefDir, familySymbol);
                                var inserted = doc.GetElement(instance.Id);
                                var depth = inserted.GetParameters("Depth").First();
                                depth.Set(wall.Width);
                                var D = inserted.GetParameters("D").First();
                                D.Set(GetDiameter(horizontalDiff, verticalDiff, wall.Width, duct.Diameter));
                            }
                        }
                    }
                }
                t.Commit();
            }
           
            return Result.Succeeded;
        }

        private double GetDiameter(double horizontalDiff, double verticalDiff, double width, double diameter)
        {
            var diff = Math.Sqrt(verticalDiff * verticalDiff + horizontalDiff * horizontalDiff);
            if(diff > 0)
            {
                var coff = diff / width;
                var edge = diameter * coff;
                var newDiameter = Math.Sqrt(diameter * diameter + edge * edge);
                return newDiameter + diff;
            }
            return diameter;
        }

        private XYZ GetRefDir(Face face)
        {
            var bboxUV = face.GetBoundingBox();
            var center = (bboxUV.Max + bboxUV.Min) / 2d;
            var normal = face.ComputeNormal(center);
            return normal.CrossProduct(XYZ.BasisZ);
        }

        private XYZ FindIntersection(Curve ductCurve, Face face)
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

        private Curve FindDuctCurve(Duct duct)
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

        private List<Face> FindWallFace(Wall wall)
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

        private FamilySymbol GetFamilySymbol(Family family)
        {
            var elementId = family.GetFamilySymbolIds().FirstOrDefault();

            if(elementId == null)
            {
                return null;
            }

            FamilySymbol familySymbol = null;
            familySymbol = family.Document.GetElement(elementId) as FamilySymbol;

            return familySymbol;
        }

        private Family FindOrLoadFamily(Document doc)
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
            return family;
        }
    }
}
