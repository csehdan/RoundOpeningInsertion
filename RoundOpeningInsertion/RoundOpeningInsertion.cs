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
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            Family family = FindOrLoadFamily(doc);

            if(family == null)
            {
                return Result.Failed;
            }
            FamilySymbol familySymbol = GetFamilySymbol(family);

            if(familySymbol == null)
            {
                return Result.Failed;
            }

            if (!familySymbol.IsActive)
            {
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Activate family symbol");
                    familySymbol.Activate();
                    tx.Commit();
                }
            }

            FilteredElementCollector walls = new FilteredElementCollector(doc).OfClass(typeof(Wall));
            using (Transaction t = new Transaction(doc, "Duct Wall Intersection"))
            {
                t.Start();

                foreach (Wall wall in walls)
                {
                    BoundingBoxXYZ bounding = wall.get_BoundingBox(null);
                    Outline outline = new Outline(bounding.Min, bounding.Max);
                    BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline);
                    FilteredElementCollector ducts = new FilteredElementCollector(doc).OfClass(typeof(Duct));
                    List<Duct> intersectedDucts = ducts.WherePasses(bbfilter).Cast<Duct>().ToList();

                    if (intersectedDucts.Count > 0)
                    {
                        List<Face> wallFaces = FindWallFace(wall);

                        if (wallFaces.Count == 2)
                        {
                            foreach (Duct duct in intersectedDucts)
                            {

                                Curve ductCurve = FindDuctCurve(duct);
                                Face face = wallFaces[0];
                                XYZ frontIntersection = FindIntersection(ductCurve, face);
                                XYZ backIntersection = FindIntersection(ductCurve, wallFaces[1]);
                                XYZ frontRefDir = GetRefDir(face);
                                XYZ intersection = null;
                                double verticalDiff = 0;
                                verticalDiff = frontIntersection.Z - backIntersection.Z;
                                double horizontalDiff = 0;

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

                                FamilyInstance instance = doc.Create.NewFamilyInstance(face, intersection, frontRefDir, familySymbol);
                                Element inserted = doc.GetElement(instance.Id);
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
            double diff = Math.Sqrt(verticalDiff * verticalDiff + horizontalDiff * horizontalDiff);
            if(diff > 0)
            {
                double coff = diff / width;
                double edge = diameter * coff;
                double newDiameter = Math.Sqrt(diameter * diameter + edge * edge);
                return newDiameter + diff;
            }
            return diameter;
        }

        private XYZ GetRefDir(Face face)
        {
            BoundingBoxUV bboxUV = face.GetBoundingBox();
            UV center = (bboxUV.Max + bboxUV.Min) / 2.0;
            XYZ location = face.Evaluate(center);
            XYZ normal = face.ComputeNormal(center);
            return normal.CrossProduct(XYZ.BasisZ);

        }

        private XYZ FindIntersection(Curve ductCurve, Face face)
        {
            IntersectionResultArray intersectionR = new IntersectionResultArray();

            SetComparisonResult results;

            results = face.Intersect(ductCurve, out intersectionR);

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
            IList<XYZ> list = new List<XYZ>();

            ConnectorSetIterator csi = duct.ConnectorManager.Connectors.ForwardIterator();
            while (csi.MoveNext())
            {
                Connector conn = csi.Current as Connector;
                list.Add(conn.Origin);
            }
            Curve curve = Line.CreateBound(list.ElementAt(0), list.ElementAt(1)) as Curve;
            curve.MakeUnbound();

            return curve;
        }

        private List<Face> FindWallFace(Wall wall)
        {
            List<Face> normalFaces = new List<Face>();

            Options opt = new Options();
            opt.ComputeReferences = true;
            opt.DetailLevel = ViewDetailLevel.Fine;

            GeometryElement e = wall.get_Geometry(opt);

            foreach (GeometryObject obj in e)
            {
                Solid solid = obj as Solid;

                if (solid != null && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        PlanarFace pf = face as PlanarFace;
                        if (pf != null)
                        {
                            if ((int)pf.FaceNormal.Z == 0)
                            {
                                if (Math.Abs(wall.Orientation.X) == Math.Abs(pf.FaceNormal.X) || Math.Abs(wall.Orientation.Y) == Math.Abs(pf.FaceNormal.Y))
                                {
                                    normalFaces.Add(pf);
                                }
                            }
                        }
                    }
                }
            }
            return normalFaces;
        }

        private FamilySymbol GetFamilySymbol(Family family)
        {
            ElementId elementId = family.GetFamilySymbolIds().FirstOrDefault();

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
            FilteredElementCollector a = new FilteredElementCollector(doc).OfClass(typeof(Family));
            Family family = a.FirstOrDefault(e => e.Name.Equals(familyName)) as Family;

            if (null == family)
            {
                string workingDirectory = Environment.CurrentDirectory;
                string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
                string FamilyPath = Path.Combine(projectDirectory, familyName + ".rfa");

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
