#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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

      var links = new FilteredElementCollector(doc)
          .OfClass(typeof(RevitLinkInstance))
          .ToList();

      var loadedExternalFilesRef = new List<RevitLinkType>();

      var forms = new FilteredElementCollector(doc)
          .WhereElementIsNotElementType()
          .OfCategory(BuiltInCategory.OST_Mass)
          .Where(f => ((FamilyInstance)f).Symbol.LookupParameter("Группа модели").AsString() != string.Empty)
          .ToList();

      var sectionForms = new Dictionary<string, BoundingBoxIsInsideFilter>();

      foreach (FamilyInstance form in forms)
      {
        var sectionName = form.Symbol.LookupParameter("Группа модели")?.AsString();
        if (sectionName == null) continue;
        //        var formSolid = GetSolidFromElement(form);
        var bbForm = form.get_BoundingBox(null);

        var formOutline = new Outline(bbForm.Min, bbForm.Max);

        var formFilter = new BoundingBoxIsInsideFilter(formOutline);

        sectionForms.Add(sectionName, formFilter);


        //set BS_Block to current document elements
        var allLinkModelElementsByForm = GetModelElementsByForm(doc, formFilter);

        using (Transaction tx = new Transaction(doc))
        {
          tx.Start("Assign section");
          foreach (var linkElement in allLinkModelElementsByForm)
          {

            var par = linkElement.LookupParameter("BS_Блок");
            if (par != null)
            {
              if (!par.IsReadOnly)
              {
                par.Set(sectionName);
              }
            }
          }
          tx.Commit();

        }
      }


      foreach (RevitLinkInstance link in links)
      {
        RevitLinkType linkType = doc.GetElement(link.GetTypeId()) as RevitLinkType;

        var linkRef = linkType.GetExternalFileReference();
        if (null == linkRef || !linkType.Name.Contains("АР"))
          continue;

        if (!linkType.IsNestedLink)
        {
          loadedExternalFilesRef.Add(linkType);
          if (linkRef.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
          {
            linkType.Unload(null);
          }
        }

        OpenOptions openOpts = new OpenOptions();
        openOpts.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets));

        var currentDoc = app.OpenDocumentFile(linkRef.GetPath(), openOpts);

        foreach (var form in sectionForms)
        {
          var allLinkModelElementsByForm = GetModelElementsByForm(currentDoc, form.Value);

          using (Transaction tx = new Transaction(currentDoc))
          {
            tx.Start("Assign section");
            foreach (var linkElement in allLinkModelElementsByForm)
            {

              var par = linkElement.LookupParameter("BS_Блок");
              if (par != null)
              {
                if (!par.IsReadOnly)
                {
                  par.Set(form.Key);
                }
              }
            }
            tx.Commit();

          }
        }

        SyncWithCentral(currentDoc);
        currentDoc.Close(false);
      }

      //reload links 
      foreach (var link in loadedExternalFilesRef)
      {
        link.Load();
      }
      return Result.Succeeded;
    }

    //    private List<>

    private Solid GetSolidFromElement(Element element)
    {
      foreach (GeometryObject geometryObject in element.get_Geometry(new Options()
      {
        IncludeNonVisibleObjects = true,
        DetailLevel = ViewDetailLevel.Fine
      }).GetTransformed(Transform.Identity))
      {
        if (geometryObject is Solid solid && Math.Abs(solid.Volume) > 0.0001)
          return solid;
      }
      return null;
    }



    private List<Element> GetModelElementsByForm(Document doc, BoundingBoxIsInsideFilter bbFilter)
    {

      var binCategories = Enum.GetValues(typeof(BuiltInCategory)).Cast<int>().ToList();

      var categories = ViewSchedule.GetValidCategoriesForSchedule().AsEnumerable()
        .Where(x => binCategories.Contains(x.IntegerValue))
        .Select(x => x.IntegerValue)
        .ToList();

      var allLinkElements = new FilteredElementCollector(doc)
        .WhereElementIsNotElementType()
        .WherePasses(bbFilter)
        .Where(x => x.Category != null &&
                    x.IsValidObject &&
                    ((x.Location != null && (x.Location is LocationCurve || x.Location is LocationPoint)) || categories.Contains(x.Category.Id.IntegerValue)) &&
                    x.GetTypeId().IntegerValue > 0 &&
                    !(x is ProjectLocation) &&
                    !(x is View) &&
                    !(x is Room))
        .ToList();
      return allLinkElements;
    }

    public void SyncWithCentral(Document doc)
    {
      // set options for accessing central model
      var transOpts = new TransactWithCentralOptions();
      //      var transCallBack = new SyncLockCallback();
      // override default behavioor of waiting to try sync if central model is locked
      //      transOpts.SetLockCallback(transCallBack);
      // set options for sync with central
      var syncOpts = new SynchronizeWithCentralOptions();
      var relinquishOpts = new RelinquishOptions(true);
      syncOpts.SetRelinquishOptions(relinquishOpts);
      // do not autosave local model
      syncOpts.SaveLocalAfter = false;
      syncOpts.Comment = "Назначен BS_Блок";
      try
      {
        doc.SynchronizeWithCentral(transOpts, syncOpts);
      }
      catch (Exception ex)
      {
        Debug.Write($"Sync with model {doc.Title}", ex.Message);
      }
    }
  }
}
