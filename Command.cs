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

      var sectionForms = new Dictionary<string, BoundingBoxIntersectsFilter>();

      foreach (FamilyInstance form in forms)
      {
        var sectionName = form.Symbol.LookupParameter("Группа модели")?.AsString();
        if (sectionName == null) continue;
        //        var formSolid = GetSolidFromElement(form);
        var bbForm = form.get_BoundingBox(null);

        var formOutline = new Outline(bbForm.Min, bbForm.Max);

        var bbFilter = new BoundingBoxIntersectsFilter(formOutline);

        sectionForms.Add(sectionName, bbFilter);


        //set BS_Block to current document elements
        var modelElementsByForm = GetModelElementsByForm(doc, bbFilter);

        assingSectionToElements(doc, sectionName, modelElementsByForm);
      }


      foreach (RevitLinkInstance link in links)
      {
        RevitLinkType linkType = doc.GetElement(link.GetTypeId()) as RevitLinkType;

        var linkRef = linkType.GetExternalFileReference();
        if (null == linkRef)
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
          assingSectionToElements(currentDoc, form.Key, allLinkModelElementsByForm);
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

    private static void assingSectionToElements(Document doc, string sectionName, List<Element> modelElementsByForm)
    {
      using (Transaction tx = new Transaction(doc))
      {
        tx.Start("Assign section");
        foreach (var e in modelElementsByForm)
        {
          var par = e.LookupParameter("BS_Блок");
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

    private List<Element> GetModelElementsByForm(Document doc, BoundingBoxIntersectsFilter bbFilter)
    {

      var binCategories = Enum.GetValues(typeof(BuiltInCategory)).Cast<int>().ToList();

      var modelCategories = ViewSchedule.GetValidCategoriesForSchedule().AsEnumerable()
        .Where(x => binCategories.Contains(x.IntegerValue))
        .Select(x => x.IntegerValue)
        .ToList();

      var allModelElements = new FilteredElementCollector(doc)
        .WhereElementIsNotElementType()
        .WherePasses(bbFilter)
        .Where(x => x.Category != null &&
                    x.IsValidObject &&
                    x.Location != null  &&
                    modelCategories.Contains(x.Category.Id.IntegerValue) &&
                    x.GetTypeId().IntegerValue > 0 &&
                    x.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Mass &&
                    x.Category.Id.IntegerValue != (int)BuiltInCategory.OST_ProjectBasePoint &&
                    x.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Views
                    ).ToList();
       
      

      return allModelElements;
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
