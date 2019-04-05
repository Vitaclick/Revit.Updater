#region Namespaces
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using ProgressChangedEventArgs = Autodesk.Revit.DB.Events.ProgressChangedEventArgs;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using TaskDialogResult = Autodesk.Revit.UI.TaskDialogResult;

#endregion

namespace RevitUpdater
{
  public static class Globals
  {
    public static string data;
  }
  class AppEvents : IExternalApplication
  {
    public List<ModelPath> data;

    public Result OnStartup(UIControlledApplication a)
    {

      // Register related events
      try
      {
        // register events
        a.ControlledApplication.ProgressChanged += app_ClosingLinks;
        a.ControlledApplication.DocumentOpening += app_DocumentOpening;
        a.DialogBoxShowing += new EventHandler<Autodesk.Revit.UI.Events.DialogBoxShowingEventArgs>(AppDialogShowing);
        a.ControlledApplication.FailuresProcessing += new EventHandler<FailuresProcessingEventArgs>(OnFailuresProcessing);
//        a.ControlledApplication.FailuresProcessing += new EventHandler<FailuresProcessingEventArgs>(CheckWarnings);
      }
      catch (Exception) { return Result.Failed;}

      return Result.Succeeded;
    }
    
    public Result OnShutdown(UIControlledApplication a)
    {
      a.ControlledApplication.ProgressChanged -= app_ClosingLinks;
//      a.ControlledApplication.DocumentOpening -= app_DocumentOpening;
      a.DialogBoxShowing -= new EventHandler<Autodesk.Revit.UI.Events.DialogBoxShowingEventArgs>(AppDialogShowing);
      a.ControlledApplication.FailuresProcessing -= new EventHandler<FailuresProcessingEventArgs>(OnFailuresProcessing);
//      a.ControlledApplication.FailuresProcessing -= new EventHandler<FailuresProcessingEventArgs>(CheckWarnings);
      return Result.Succeeded;
    }

    private void CheckWarnings(object sender, FailuresProcessingEventArgs e)
    {
      FailuresAccessor fa = e.GetFailuresAccessor();

      IList<FailureMessageAccessor> failList = fa.GetFailureMessages(); // Inside event handler, get all warnings

      foreach (FailureMessageAccessor failure in failList)
      {
        // check FailureDefinitionIds against ones that you want to dismiss, 
        FailureDefinitionId failID = failure.GetFailureDefinitionId();
        var failureSeverity = fa.GetSeverity();
        if (failureSeverity == FailureSeverity.Warning)
        {
          fa.DeleteWarning(failure);
        }
//        // prevent Revit from showing Unenclosed room warnings
//        if (failID == BuiltInFailures.RoomFailures.RoomNotEnclosed)
//        {
//          fa.DeleteWarning(failure);
//        }
      }
    }

    void AppDialogShowing(object sender, DialogBoxShowingEventArgs args)
    {
      // DialogBoxShowingEventArgs has two subclasses - TaskDialogShowingEventArgs & MessageBoxShowingEventArgs
      // In this case we are interested in this event if it is TaskDialog being shown. 
      if (args is TaskDialogShowingEventArgs e)
      {
        //worry about this later - 1002 = cancel
        if (e.DialogId == "TaskDialog_Unresolved_References") {
          e.OverrideResult(1002);
        }
        //Don't sync newly created files. 1003 = close
        if (e.DialogId == "TaskDialog_Local_Changes_Not_Synchronized_With_Central") {
          e.OverrideResult(1003);
        }
        if (e.DialogId == "TaskDialog_Save_Changes_To_Local_File") {
          //Relinquish unmodified elements and worksets
          e.OverrideResult(1001);
        }
        if (e.DialogId == "TaskDialog_Copied_Central_Model")
        {
          args.OverrideResult((int)TaskDialogResult.Close);
        }
        if (e.Message == "The floor/roof overlaps the highlighted wall(s). Would you like to join geometry and cut the overlapping volume out of the wall(s)?")
        {
          // Call OverrideResult to cause the dialog to be dismissed with the specified return value
          // (int) is used to convert the enum TaskDialogResult.No to its integer value which is the data type required by OverrideResult
          e.OverrideResult((int)TaskDialogResult.No);
        }
      }

//      // Get the help id of the showing dialog
//      int dialogId = args.HelpId;
//
//      // Format the prompt information string
//      String promptInfo = "A Revit dialog will be opened.\n";
//      promptInfo += "The help id of this dialog is " + dialogId.ToString() + "\n";
//      promptInfo += "If you don't want the dialog to open, please press cancel button";
//
//      // Show the prompt message, and allow the user to close the dialog directly.
//      TaskDialog taskDialog = new TaskDialog("Revit");
//      taskDialog.MainContent = promptInfo;
//      TaskDialogCommonButtons buttons = TaskDialogCommonButtons.Ok |
//                                        TaskDialogCommonButtons.Cancel;
//      taskDialog.CommonButtons = buttons;
//      TaskDialogResult result = taskDialog.Show();
//      if (TaskDialogResult.Cancel == result)
//      {
//        // Do not show the Revit dialog
//        args.OverrideResult(1);
//      }
//      else
//      {
//        // Continue to show the Revit dialog
//        args.OverrideResult(0);
//      }
    }

    private void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
    {
      FailuresAccessor failuresAccessor = e.GetFailuresAccessor();
      IEnumerable<FailureMessageAccessor> failureMessages = failuresAccessor.GetFailureMessages();
      foreach (FailureMessageAccessor failureMessage in failureMessages) {
        if (failureMessage.GetSeverity() == FailureSeverity.Warning) {
          failuresAccessor.DeleteWarning(failureMessage);
        }
      }
      e.SetProcessingResult(FailureProcessingResult.Continue);
    }

    private void app_DocumentOpening(object sender, DocumentOpeningEventArgs e)
    {
//      try
//      {
//        var temp = new List<ModelPath>();
//
//        var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(@"RSN://" + e.PathName);


//        var trans = new TransmissionData(TransmissionData.ReadTransmissionData(modelPath));
//        foreach (var refId in trans.GetAllExternalFileReferenceIds())
//        {
//          var extRef = trans.GetLastSavedReferenceData(refId);
//
//          if (extRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
//          {
//            var path = extRef.GetPath();
//            temp.Add(path);
//          }
//        }
        Globals.data = Path.GetFileNameWithoutExtension(e.PathName);

//      }
//        catch (Exception exception)
//      {
//        Console.WriteLine(exception);
//        throw;
//      }

    }


    public void app_ClosingLinks(object sender, ProgressChangedEventArgs e)
    {
      if (Globals.data != null)
      {
//        var openingLinksNames = new List<string>();
//        foreach (ModelPath modelPath in Globals.data)
//        {
//          var filePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
//          openingLinksNames.Add(Path.GetFileNameWithoutExtension(filePath));
//        }
        // get document from event args
//        if (openingLinksNames.Any(file => e.Caption.Contains(file)))
        var state = e.Caption;
        if ( (state.StartsWith("Loading") || state.StartsWith("Загрузка")) && !state.Contains(Globals.data) )
        {
//          if (file.Length > 0)
//          {
//            if (!state.Contains(Globals.data))
//            {
              if (e.Cancellable)
              {
                Debug.Write(state);
                e.Cancel();
              }
//            }
//          }
        }
      }
    }
  }
}
