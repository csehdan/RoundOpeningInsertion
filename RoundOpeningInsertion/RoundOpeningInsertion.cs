using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
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
            var family = Helpers.FindOrLoadFamily(doc, familyName);
            var familySymbol = Helpers.ResolveFamilySymbol(family, doc);

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
                        var wallFaces = Helpers.FindWallFace(wall).ToList();

                        if (wallFaces.Count == 2)
                        {
                            foreach (var duct in intersectedDucts)
                            {
                                var ductCurve = Helpers.FindDuctCurve(duct);
                                var face = wallFaces[0];
                                var frontIntersection = Helpers.FindIntersection(ductCurve, face);
                                var backIntersection = Helpers.FindIntersection(ductCurve, wallFaces[1]);

                                if (frontIntersection is null || backIntersection is null)
                                {
                                    continue;
                                }

                                var frontRefDir = Helpers.GetRefDir(face);
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
                                D.Set(Helpers.GetDiameter(horizontalDiff, verticalDiff, wall.Width, duct.Diameter));
                            }
                        }
                    }
                }
                t.Commit();
            }
           
            return Result.Succeeded;
        }
    }
}
