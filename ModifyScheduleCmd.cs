//using System;
//using System.Collections.Generic;
//using System.Linq;
//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;
//using Autodesk.Revit.UI.Selection;

//namespace LINE.Revit
//{
//    [Transaction(TransactionMode.Manual)]
//    public class ModifyScheduleCmd : IExternalCommand
//    {
//        private readonly Guid schemaGUID = new Guid("91c053bd-edeb-4feb-abac-ef862c311e9d");
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            try
//            {
//                // Make sure that we're in a schedule view or that a schedule is selected
//                View activeView = commandData.Application.ActiveUIDocument.Document.ActiveView;
//                Selection selection = commandData.Application.ActiveUIDocument.Selection;
//                if (selection.GetElementIds().Count == 1)
//                {
//                    try
//                    {
//                        ElementId selectedId = selection.GetElementIds().FirstOrDefault();
//                        Element selectedElem = commandData.Application.ActiveUIDocument.Document.GetElement(selectedId);
//                        if (selectedElem.Category.Id.IntegerValue == commandData.Application.ActiveUIDocument.Document.Settings.Categories.get_Item(BuiltInCategory.OST_ScheduleGraphics).Id.IntegerValue)
//                        {
//                            FilteredElementCollector scheduleCollector = new FilteredElementCollector(commandData.Application.ActiveUIDocument.Document);
//                            scheduleCollector.OfClass(typeof(ViewSchedule));
//                            foreach (ViewSchedule vs in scheduleCollector)
//                            {
//                                if (vs.Name == selectedElem.Name)
//                                {
//                                    activeView = vs as View;
//                                }
//                            }
//                        }
//                    }
//                    catch { }
//                }

//                if (activeView.ViewType == ViewType.Schedule)
//                {
//                    ViewSchedule vs = activeView as ViewSchedule;
//                    int index = -1;

//                    // Get the schema
//                    Autodesk.Revit.DB.ExtensibleStorage.Schema schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(schemaGUID);
//                    if (schema != null)
//                    {
//                        Autodesk.Revit.DB.ExtensibleStorage.Entity entity = null;
//                        try
//                        {
//                            entity = SchemaManager.ReadEntityData(commandData.Application.ActiveUIDocument.Document, schema);
//                        }
//                        catch { }

//                        if (entity != null)
//                        {
//                            IList<ElementId> elementIds = entity.Get<IList<ElementId>>("ScheduleId");
//                            IList<string> paths = entity.Get<IList<string>>("ExcelFilePath");
//                            IList<string> worksheets = entity.Get<IList<string>>("WorksheetName");
//                            IList<string> dateTimes = entity.Get<IList<string>>("DateTime");
//                            IList<int> pathTypes;
//                            try
//                            {
//                                pathTypes = entity.Get<IList<int>>("PathType");
//                            }
//                            catch
//                            {
//                                List<int> tempPaths = new List<int>();
//                                for (int i = 0; i < elementIds.Count; i++)
//                                {
//                                    tempPaths.Add(0);
//                                }
//                                pathTypes = tempPaths;
//                            }

//                            for (int i = 0; i < elementIds.Count; i++)
//                            {
//                                if (elementIds[i].IntegerValue == vs.Id.IntegerValue)
//                                    index = i;
//                            }

//                            if (index >= 0)
//                            {
//                                Scheduler scheduler = new Scheduler();
//                                scheduler.ModifySchedule(commandData.Application.ActiveUIDocument.Document, elementIds[index], paths[index], worksheets[index], "Update Excel Schedule", pathTypes[index], false);
//                            }
//                            else
//                            {
//                                TaskDialog.Show("Error", "Could not find matching Excel data for this schedule.");
//                            }
//                        }
//                        else
//                        {
//                            TaskDialog.Show("Error", "Could not find matching Excel data for this schedule.");
//                        }
//                    }
//                    else
//                    {
//                        TaskDialog.Show("Error", "Could not find matching Excel data for this schedule.");
//                    }
//                }
//                else
//                {
//                    TaskDialog.Show("Error", "This command only works when the active view is a Schedule, or if a selected schedule, was created from the import");
//                }

//                // Write to home
//                RevitCommon.FileUtils.WriteToHome("Excel Import - Modify Link", commandData.Application.Application.VersionName, commandData.Application.Application.Username);

//                return Result.Succeeded;
//            }
//            catch (Exception ex)
//            {
//                message = ex.Message;
//                return Result.Failed;
//            }
//        }
//    }
//}
