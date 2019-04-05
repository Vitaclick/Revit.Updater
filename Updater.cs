#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Application = Autodesk.Revit.ApplicationServices.Application;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
#endregion

namespace RevitUpdater
{
  [Transaction(TransactionMode.Manual)]
  public class Updater : IExternalCommand
  {
    public static class Globals
    {
      public static List<List<ModelPath>> data = new List<List<ModelPath>>(){};
    }

    private void app_ListOfLinks(ModelPath modelPath)
    {
      var temp = new List<ModelPath>();

      var trans = new TransmissionData(TransmissionData.ReadTransmissionData(modelPath));
      foreach (var refId in trans.GetAllExternalFileReferenceIds())
      {
        var extRef = trans.GetLastSavedReferenceData(refId);

        if (extRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
        {
          var path = extRef.GetPath();
          temp.Add(path);
        }
      }

      Globals.data.Add(temp);
    }

//    public void app_ClosingLinks(object sender, ProgressChangedEventArgs e)
//    {
//      if (RevitUpdater.Globals.data != null)
//      {
//        var openingLinksNames = new List<string>();
//        foreach (ModelPath modelPath in RevitUpdater.Globals.data)
//        {
//          var filePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
//          openingLinksNames.Add(Path.GetFileNameWithoutExtension(filePath));
//        }
//        // get document from event args
//        if (openingLinksNames.Any(file => e.Caption.Contains(file)) && e.LowerRange == e.Position)
//        {
//          if (e.Cancellable)
//          {
//            Debug.Write(e.Caption);
//
//            e.Cancel();
//          }
//        }
//      }
//    }


        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            
//            uiapp.Application.ProgressChanged += new EventHandler<ProgressChangedEventArgs>(app_ClosingLinks);
//            uiapp.Application.DocumentOpening += new EventHandler<DocumentOpeningEventArgs>(app_DocumentOpening);


            var modelPaths = GetModelPaths();
            if (modelPaths == null) return Result.Cancelled;

            // Set open configuration all-in
            var openOpts = new OpenOptions();
            openOpts.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets));

          var preDocument = uidoc;

          foreach (var modelPath in modelPaths)
          {
            app_ListOfLinks(modelPath);
            // ver 2 opendactivedocument
//            var openingUI = uiapp.OpenAndActivateDocument(modelPath, openOpts, false);
//            var openedDoc = openingUI.Document;
//
//            if (preDocument != null)
//            {
//              preDocument = openingUI;
//            }


            // ver 1 opendocumentfile
//            var openedDoc = app.OpenDocumentFile(modelPath, openOpts);
//            var trans = new Transaction(openedDoc, "kek");
//            trans.Start();
//            openedDoc.Regenerate();
//            trans.Commit();

//            SyncWithCentral(openedDoc);
//            openedDoc.Close(false);
          }

          //            CloseModelLinks(testModel);
          // Open and sync model
          //            var openedDoc = app.OpenDocumentFile(testModel, openOpts);

//                uiapp.Application.ProgressChanged -= app_ClosingLinks;
//      uiapp.Application.DocumentOpening -= app_DocumentOpening;

            return Result.Succeeded;

        }

        private List<ModelPath> GetModelPaths()
        {
            // Collect folders to update
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result.ToString() == "Ok")
            {
                var path = dialog.FileName;
                var directoryInfo = new DirectoryInfo(path);

                return ParseModelPaths(directoryInfo);
            }
            else
            {
                return null;
            }

        }

        private void CloseModelLinks(ModelPath modelPath)
        {
            // Access transmission data in the given file
            TransmissionData trans = TransmissionData.ReadTransmissionData(modelPath);
            if (trans != null)
            {
                // collect all external references
                var externalReferences = trans.GetAllExternalFileReferenceIds();
                foreach (ElementId extRefId in externalReferences)
                {
                    var extRef = trans.GetLastSavedReferenceData(extRefId);
                    if (extRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
                    {
//                        var p = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetPath());
                        //set data
                        trans.SetDesiredReferenceData(extRefId, extRef.GetPath(), extRef.PathType, false);
                        Debug.Write($"{extRef.GetPath().CentralServerPath} " +
                                    $"{extRef.GetPath().ServerPath}");
                    }

                }
                // make sure the IsTransmitted property is set
                trans.IsTransmitted = true;

                // modified transmissoin data must be saved back to the model
                TransmissionData.WriteTransmissionData(modelPath, trans);
            }
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
            syncOpts.Comment = "Освобождено";
            try
            {
                doc.SynchronizeWithCentral(transOpts, syncOpts);
            }
            catch (Exception ex)
            {
                TaskDialog.Show($"Sync with model {doc.Title}", ex.Message);
            }



        }
        private List<ModelPath> ParseModelPaths(DirectoryInfo path)
        {
            var revitFiles = path.GetDirectories("*.rvt", System.IO.SearchOption.AllDirectories);
            var modelPaths = new List<ModelPath>();

            // Parse filepath to RSN path
            foreach (var filePathInfo in revitFiles)
            {
                var filePath = filePathInfo.FullName;
                var rsnPath = Regex.Replace(@"RSN:\\" + filePath.Remove(0, 2), @"(?<=.ru\b).*?(?<=Prj|Projects|Prg)", "");
                modelPaths.Add(ModelPathUtils.ConvertUserVisiblePathToModelPath(rsnPath));
            }
            return modelPaths;

        }

    }
}
