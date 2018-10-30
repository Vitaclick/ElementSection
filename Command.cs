#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace ElementSection
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Access current selection

            Selection sel = uidoc.Selection;

            // Retrieve elements from database

            FilteredElementCollector col
                = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Mass);

            Options opt = new Options();
            opt.ComputeReferences = true;
            opt.DetailLevel = ViewDetailLevel.Fine;

            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .ToList();



            foreach (FamilyInstance form in col.ToList())
            {
                ICollection<Element> sectionElements = new List<Element>();

                var sectionName = form.Symbol.LookupParameter("Группа модели").AsString();

                var sectionBB = form.get_BoundingBox(null);
                var sectionGeometry = GetSolidFromElement(form);

                var bbFilter = GetBBFilter(sectionBB);

                var intersectionResult = new IntersectionResultArray();

                foreach (RevitLinkInstance link in links)
                {

                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Transaction Name");
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc != null)
                        {

                            var allLinkElements = new FilteredElementCollector(linkDoc)
                                .WhereElementIsNotElementType()
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .ToList();

                            foreach (var linkElement in allLinkElements)
                            {

                                Solid linkElementGeometry = GetSolidFromElement(linkElement);

                                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                                    linkElementGeometry, sectionGeometry, BooleanOperationsType.Intersect);
                                double intersectionVolume = intersection.Volume;
                                if (intersectionVolume <= 0) continue;

                                // set parameter to linked element
                                //                                sectionElements.Add(linkElement);
                                var par = linkElement.LookupParameter("BS_Номер");
                                par.Set(sectionName);

                            }
                        }
                        tx.Commit();

                    }



                }



            }

            return Result.Succeeded;
        }

        private BoundingBoxIntersectsFilter GetBBFilter(BoundingBoxXYZ bb)
        {

            // Диагональ формообразующей
            Outline outline = new Outline(
                new XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z));

            return new BoundingBoxIntersectsFilter(outline);
        }

        public Solid GetSolidFromElement(Element element)
        {
            Solid solid = null;
            Options opt = new Options();
            GeometryElement geomElem = element.get_Geometry(opt);
            if (geomElem == null)
            {
                return null;
            }

            foreach (GeometryObject geomObj in geomElem)
            {
                solid = geomObj as Solid;
                if (solid != null && solid.Volume > 0)
                {
                    return solid;
                }
            }

            return solid;
        }
    }
}
