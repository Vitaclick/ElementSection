using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElementSection
{
  [Transaction(TransactionMode.Manual)]
  public class Search : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {

      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = uidoc.Document;

      var links = new FilteredElementCollector(doc)
        .OfClass(typeof(RevitLinkInstance))
        .ToList();
      var loadedExternalFilesRef = new List<RevitLinkType>();

      var o = new List<string>();
      foreach (RevitLinkInstance link in links)
      {
        RevitLinkType linkType = doc.GetElement(link.GetTypeId()) as RevitLinkType;
        var linkRef = linkType.GetExternalFileReference();
        if (null == linkRef || !linkType.Name.Contains("АР"))
          continue;

        var currentDoc = link.Document;

        var binCategories = Enum.GetValues(typeof(BuiltInCategory)).Cast<int>().ToList();

        var categories = ViewSchedule.GetValidCategoriesForSchedule().AsEnumerable()
          .Where(x => binCategories.Contains(x.IntegerValue))
          .Select(x => x.IntegerValue)
          .ToList();

        var allLinkElements = new FilteredElementCollector(currentDoc)
          .WhereElementIsNotElementType()
          .Where(x => x.Category != null &&
                      x.IsValidObject &&
                      ((x.Location != null && (x.Location is LocationCurve || x.Location is LocationPoint)) ||
                       categories.Contains(x.Category.Id.IntegerValue)) &&
                      x.GetTypeId().IntegerValue > 0 &&
                      !(x is ProjectLocation) &&
                      !(x is View))
          .ToList();


        foreach (Element e in allLinkElements)
        {
          if (e.LookupParameter("BS_Блок").AsString() == "Секция 05")
          {
            o.Add(e.ToString());
          }
        }

        TaskDialog.Show("kek", o.ToString());
      }

      return Result.Succeeded;

    }
  }
}
