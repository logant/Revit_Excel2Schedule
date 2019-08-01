#region Header
//
//Copyright(c) 2019 Timothy Logan, HKS Inc

//Permission is hereby granted, free of charge, to any person obtaining
//a copy of this software and associated documentation files (the
//"Software"), to deal in the Software without restriction, including
//without limitation the rights to use, copy, modify, merge, publish,
//distribute, sublicense, and/or sell copies of the Software, and to
//permit persons to whom the Software is furnished to do so, subject to
//the following conditions:

//The above copyright notice and this permission notice shall be
//included in all copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
//EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
//MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
//NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
//LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
//OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
//WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
#endregion

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using RevitCommon.Attributes;
using RevitCommon.ElementExtensions;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace LINE.Revit
{
    [ExtApp(Name = "Import Excel as Schedule", Description = "Import an Excel file as a Revit schedule",
            Guid = "8b3ad4c9-76d6-4c92-9e3c-5cc0a4e058b3", Vendor = "HKSL", VendorDescription = "HKS LINE, www.hksline.com",
            ForceEnabled = false, Commands = new[] { "Import Excel as Schedule", "Manage Excel Links", "Import Excel Settings", "Excel Link Auto-Sync" })]
    public class UpdateExcelApp : IExternalApplication
    {
        
        private readonly Guid schemaGUID = (typeof(ExcelScheduleEntity).GetCustomAttributes(typeof(SchemaAttribute), true).FirstOrDefault() as SchemaAttribute)?.GUID ?? Guid.Empty;
        private readonly string dsName = Properties.Settings.Default.DataStorageName;

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        static extern int SetWindowText(IntPtr hWnd, string lpString);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        public static int Version;

        public static IntPtr RevitHandle;


        public Result OnShutdown(UIControlledApplication application)
        {
            try
            {
                application.ControlledApplication.DocumentOpened -= DocumentOpened;
            }
            catch { }
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Get the revit version
                Version = Convert.ToInt32(application.ControlledApplication.VersionNumber);

                // Get the revit handle
                RevitHandle = IntPtr.Zero;
                if (Version < 2019)
                    RevitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                else
                    RevitHandle = application.GetType().GetProperty("MainWindowHandle") != null
                        ? (IntPtr)application.GetType().GetProperty("MainWindowHandle")?.GetValue(application)
                        : IntPtr.Zero;

                application.ControlledApplication.DocumentOpened += DocumentOpened;

                ScheduleDeleteUpdater excelSchedUpdater = new ScheduleDeleteUpdater(application.ActiveAddInId);
                UpdaterRegistry.RegisterUpdater(excelSchedUpdater, true);
                //ElementCategoryFilter ecf = new ElementCategoryFilter(BuiltInCategory.OST_Views);
                ElementClassFilter ecf = new ElementClassFilter(typeof(ViewSchedule), false);
                UpdaterRegistry.AddTrigger(excelSchedUpdater.GetUpdaterId(), ecf, Element.GetChangeTypeElementDeletion());

                // create the buttons
                string path = typeof(UpdateExcelApp).Assembly.Location;

                // pushbutton for the import command
                PushButtonData excelImportPushButtonData = new PushButtonData(
                    "Import Excel As Schedule", "Import Excel", path, typeof(ImportExcelCmd).FullName)
                {
                    LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.ExcelScheduleIcon.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions()),
                    Image = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.ExcelScheduleIcon_16x16.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions()),
                    ToolTip = "Import data from an Excel worksheet into Revit as a schedule.",
                };


                // Pushbutton for the settings command
                PushButtonData settingsPushButtonData = new PushButtonData(
                    "Import Lineweights", "Settings", path, typeof(SettingsCmd).FullName)
                {
                    ToolTip = "Set the line styles to use for different lineweight settings.",
                };

                // PushButtonData for the Manage Links window
                PushButtonData manageLinksPushButtonData = new PushButtonData(
                    "Manage Links", "Manage Links", path, typeof(ManageLinksCmd).FullName)
                {
                    ToolTip = "Manage linked excel files",
                };

                SplitButtonData excelSBD = new SplitButtonData("Import Excel", "Import\nExcel");

                // Set default config values
                string helpPath = Path.Combine(Path.GetDirectoryName(typeof(UpdateExcelApp).Assembly.Location), "help\\ImportExcel.pdf");
                string tabName = "Add-Ins";
                string panelName = "Views";
                if(RevitCommon.FileUtils.GetPluginSettings(typeof(UpdateExcelApp).Assembly.GetName().Name, out Dictionary<string, string> settings))
                {
                    // Settings retrieved, lets try to use them.
                    if (settings.ContainsKey("help-path") && !string.IsNullOrWhiteSpace(settings["help-path"]))
                    {
                        // Check to see if it's relative path
                        string hp = Path.Combine(Path.GetDirectoryName(typeof(UpdateExcelApp).Assembly.Location), settings["help-path"]);
                        if (File.Exists(hp))
                            helpPath = hp;
                        else
                            helpPath = settings["help-path"];
                    }
                    if (settings.ContainsKey("tab-name") && !string.IsNullOrWhiteSpace(settings["tab-name"]))
                        tabName = settings["tab-name"];
                    if (settings.ContainsKey("panel-name") && !string.IsNullOrWhiteSpace(settings["panel-name"]))
                        panelName = settings["panel-name"];
                }
                
                // Create the SplitButton.
                SplitButton excelSB = RevitCommon.UI.AddToRibbon(application, tabName, panelName, excelSBD);

                // Setup the help
                ContextualHelp help = null;
                if (File.Exists(helpPath))
                    help = new ContextualHelp(ContextualHelpType.ChmFile, helpPath);
                else if (Uri.TryCreate(helpPath, UriKind.Absolute, out Uri uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps))
                    help = new ContextualHelp(ContextualHelpType.Url, helpPath);
                if (help != null)
                {
                    excelImportPushButtonData.SetContextualHelp(help);
                    manageLinksPushButtonData.SetContextualHelp(help);
                    settingsPushButtonData.SetContextualHelp(help);
                    excelSBD.SetContextualHelp(help);
                    excelSB.SetContextualHelp(help);
                }

                PushButton importPB = excelSB.AddPushButton(excelImportPushButtonData) as PushButton;
                PushButton managePB = excelSB.AddPushButton(manageLinksPushButtonData) as PushButton;
                PushButton settingPB = excelSB.AddPushButton(settingsPushButtonData) as PushButton;
                excelSB.IsSynchronizedWithCurrentItem = false;

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }

        }

        public void DocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            Document doc = e.Document;

            if (doc.IsFamilyDocument)
                return;

            Schema schema = Schema.Lookup(schemaGUID);
            if (schema == null || !schema.IsValidObject)
                return;

            // Check to see if there is out-dated data stored in ProjectInformation
            Entity entity = null;
            entity = doc.ProjectInformation.GetEntity(schema);
            if (entity != null && entity.IsValid())
            {
                // Need to transition the data to a datastorage object.
                // First make sure this isn't a workshared document with the ProjectInfo already checked out by another user
                // If it's checked out by another person, we'll just skip this since we can't fix it now.
                if (doc.IsWorkshared && WorksharingUtils.GetCheckoutStatus(doc, doc.ProjectInformation.Id) == CheckoutStatus.OwnedByOtherUser)
                    return;

                // Otherwise, lets transition the data from the old to the new.
                if (entity.Get<IList<ElementId>>("ScheduleId") != null)
                {
                    // Get the information from the ProjectInformation entity
                    var schedIds = entity.Get<IList<ElementId>>("ScheduleId").ToList();
                    var paths = entity.Get<IList<string>>("ExcelFilePath").ToList();
                    var wsNames = entity.Get<IList<string>>("WorksheetName").ToList();
                    var dts = entity.Get<IList<string>>("DateTime").ToList();
                    var pTypes = entity.Get<IList<int>>("PathType")?.ToList() ?? new List<int>();

                    // Purge the old Schema and Entity, then assign the data to a new Schema and DataStorage element
                    RebuildSchema(doc, schema, schedIds, paths, wsNames, dts, pTypes);
                }
            }

            // Find if a datstorage element exists now and update as needed.
            DataStorage ds = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).Where(x => x.Name.Equals(dsName)).Cast<DataStorage>().FirstOrDefault();

            // Get the ExcelScheduleEntity from the data storage and verify its valid
            ExcelScheduleEntity ent = ds?.GetEntity<ExcelScheduleEntity>();
            if (ent == null)
                return;

            // Check if any schedules need to be updated
            List<int> modifyIndices = new List<int>();
            List<DateTime> modDateTimes = new List<DateTime>();
            for (int i = 0; i < ent.ScheduleId.Count; i++)
            {
                string currentFilePath;
                string docPath;
                if (doc.IsWorkshared)
                    docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                else
                    docPath = doc.PathName;

                if ((PathType)ent.PathType[i] == PathType.Absolute)
                    currentFilePath = ent.ExcelFilePath[i];
                else
                    currentFilePath = PathExchange.GetFullPath(ent.ExcelFilePath[i], docPath);

                // Get the file write time as UTC
                DateTime modTime = new FileInfo(currentFilePath).LastWriteTimeUtc;
                DateTime storedTime = Convert.ToDateTime(ent.DateTime[i]);

                // Make sure the save time isn't more or less the same as stored.
                if ((modTime - storedTime).Seconds > 1)
                {
                    modifyIndices.Add(i);
                    modDateTimes.Add(modTime);
                }
            }

            if (modifyIndices.Count == modDateTimes.Count && modifyIndices.Count > 0)
            {
                IntPtr statusBar = FindWindowEx(RevitHandle, IntPtr.Zero, "msctls_statusbar32", "");
                foreach (int i in modifyIndices)
                {
                    if (statusBar != IntPtr.Zero)
                    {
                        SetWindowText(statusBar, string.Format("Updating Excel Schedule {0}.", ent.WorksheetName[modifyIndices[i]]));
                    }
                    Scheduler scheduler = new Scheduler();
                    scheduler.ModifySchedule(doc, ent.ScheduleId[modifyIndices[i]], ent.ExcelFilePath[modifyIndices[i]],
                        ent.WorksheetName[modifyIndices[i]], "Update Excel Schedule", ent.PathType[modifyIndices[i]], 
                        Properties.Settings.Default.reloadValuesOnly);

                    ent.DateTime[modifyIndices[i]] = modDateTimes[i].ToString();
                }
                if (statusBar != IntPtr.Zero)
                {
                    SetWindowText(statusBar, "");
                }

                // change the dateTimes
                using (Transaction t = new Transaction(doc, "Update schedule date"))
                {
                    t.Start();
                    ds.SetEntity(ent);
                    t.Commit();
                }

                // Write to home
                RevitCommon.FileUtils.WriteToHome("Excel Import - Document Open Reload", doc.Application.VersionName, doc.Application.Username);
            }
        }

        private void RebuildSchema(Document doc, Schema schema, List<ElementId> elementIds, List<string> paths, List<string> worksheets, List<string> dateTimes, List<int> pathTypes)
        {
            using (Transaction trans = new Transaction(doc, "Updating Excel Link Information"))
            {
                trans.Start();
                try
                {
                    SubTransaction deleteTrans = new SubTransaction(doc);
                    deleteTrans.Start();
                    // Delete the schema/entity from the ProjectInformation
                    doc.ProjectInformation.DeleteEntity(schema);
                    Schema.EraseSchemaAndAllEntities(schema, false);
                    deleteTrans.Commit();


                    // Start a subtransaction for create the datastorage and entity
                    SubTransaction createTrans = new SubTransaction(doc);
                    createTrans.Start();

                    // Build the schedule Entity
                    ExcelScheduleEntity schedEntity = new ExcelScheduleEntity()
                    {
                        DateTime = dateTimes,
                        ExcelFilePath = paths,
                        PathType = pathTypes,
                        ScheduleId = elementIds,
                        WorksheetName = worksheets
                    };

                    // Create a DataStorage for the entity
                    DataStorage ds = DataStorage.Create(doc);
                    ds.Name = Properties.Settings.Default.DataStorageName;
                    ds.SetEntity(schedEntity);

                    // complete the Create subtransaction
                    createTrans.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Test", ex.ToString());
                }
                trans.Commit();
            }
        }
    }
}
