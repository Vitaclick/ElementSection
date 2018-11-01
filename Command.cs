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

      var links = new FilteredElementCollector(doc)
          .OfClass(typeof(RevitLinkInstance))
          .ToList();

      var bsBlockValue = "Детский сад 20";

      var loadedExternalFilesRef = new List<RevitLinkType>();

      foreach (RevitLinkInstance link in links)
      {
        RevitLinkType linkType = doc.GetElement(link.GetTypeId()) as RevitLinkType;
//        if (linkType.IsNestedLink)
//          continue;
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

        var linkModelPath = linkRef.GetPath();

        // treow an exception if unloaded
//        var linkDoc = link.GetLinkDocument();
        // throw an exception if nested
//        var linkModelPath = linkDoc.GetWorksharingCentralModelPath();

        OpenOptions openOpts = new OpenOptions();
        openOpts.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets));

        // open link doc
//        var currentUiDoc = uiapp.OpenAndActivateDocument(linkModelPath, openOpts, false);
//        var currentDoc = currentUiDoc.Document;
        var currentDoc = app.OpenDocumentFile(linkModelPath, openOpts);

        var allLinkElements = new FilteredElementCollector(currentDoc)
          .WhereElementIsNotElementType()
          .Where(x => x.Category != null && x.IsValidObject)
          .ToList();

        using (Transaction tx = new Transaction(currentDoc))
        {
          tx.Start("Assign section");
          foreach (var linkElement in allLinkElements)
          {
            var par = linkElement.LookupParameter("BS_Блок");
            if (par != null)
            {
              if (!par.IsReadOnly)
              {
                par.Set(bsBlockValue);
              }
            }
          }
          tx.Commit();

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
        TaskDialog.Show($"Sync with model {doc.Title}", ex.Message);
      }
    }
  }
}
