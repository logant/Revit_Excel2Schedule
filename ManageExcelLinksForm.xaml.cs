using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using Autodesk.Revit.DB;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Autodesk.Revit.DB.ExtensibleStorage;
using RevitCommon.ElementExtensions;


namespace LINE.Revit
{
    /// <summary>
    /// Interaction logic for ManageExcelLinksForm.xaml
    /// </summary>
    public partial class ManageExcelLinksForm : Window
    {
        Document doc;
        Guid schemaGuid;

        IList<ElementId> elementIds  = new List<ElementId>();
        IList<string> paths = new List<string>();
        IList<string> worksheets = new List<string>();
        IList<string> dateTimes = new List<string>();
        IList<int> pathTypes = new List<int>();

        WorksheetObject worksheetObj = null;

        List<LinkData> LinkedData = null;
        List<string> PathTypes = new List<string> { "Absolute", "Relative" };

        bool pathChanged = false;

        bool contentOnly = Properties.Settings.Default.reloadValuesOnly;

        public WorksheetObject Worksheet
        {
            get { return worksheetObj; }
            set { worksheetObj = value; }
        }

        public ManageExcelLinksForm(Document _doc, Guid _schemaGuid)
        {
            doc = _doc;
            schemaGuid = _schemaGuid;
            InitializeComponent();
            
            
            // Read the schema information
            Schema schema = Schema.Lookup(schemaGuid);
            if (schema != null)
            {
                Autodesk.Revit.DB.ExtensibleStorage.DataStorage ds = null;
                ICollection<ElementId> dsCollector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).ToElementIds();
                foreach (ElementId eid in dsCollector)
                {
                    DataStorage dStor = doc.GetElement(eid) as DataStorage;
                    if (dStor.Name == Properties.Settings.Default.DataStorageName)
                    {
                        ds = dStor;
                        break;
                    }
                }

                if (ds != null)
                {
                    Autodesk.Revit.DB.ExtensibleStorage.Entity entity = null;
                    entity = ds.GetEntity(schema);
                    ExcelScheduleEntity schedEntity = ds.GetEntity<ExcelScheduleEntity>();
                    if (schedEntity != null)
                    {
                        elementIds = schedEntity.ScheduleId;
                        paths = schedEntity.ExcelFilePath;
                        worksheets = schedEntity.WorksheetName;
                        dateTimes = schedEntity.DateTime;
                        pathTypes = schedEntity.PathType;
                    }

                    //if (entity.IsValid())
                    //{
                    //    elementIds = entity.Get<IList<ElementId>>("ScheduleId");
                    //    paths = entity.Get<IList<string>>("ExcelFilePath");
                    //    worksheets = entity.Get<IList<string>>("WorksheetName");
                    //    dateTimes = entity.Get<IList<string>>("DateTime");
                    //    try
                    //    {
                    //        pathTypes = entity.Get<IList<int>>("PathType");
                    //    }
                    //    catch
                    //    {
                    //        List<int> tempPaths = new List<int>();
                    //        for (int i = 0; i < elementIds.Count; i++)
                    //        {
                    //            tempPaths.Add(0);
                    //        }
                    //        pathTypes = tempPaths;
                    //    }
                    //}
                }
                   

                ////Autodesk.Revit.DB.ExtensibleStorage.Entity entity = null;
                //try
                //{
                //    entity = doc.ProjectInformation.GetEntity(schema);
                //}
                //catch { }

                //if (entity != null)
                //{
                //    try
                //    {
                //        elementIds = entity.Get<IList<ElementId>>("ScheduleId");
                //        paths = entity.Get<IList<string>>("ExcelFilePath");
                //        worksheets = entity.Get<IList<string>>("WorksheetName");
                //        dateTimes = entity.Get<IList<string>>("DateTime");
                //        try
                //        {
                //            pathTypes = entity.Get<IList<int>>("PathType");
                //        }
                //        catch
                //        {
                //            List<int> tempPaths = new List<int>();
                //            for (int i = 0; i < elementIds.Count; i++)
                //            {
                //                tempPaths.Add(0);
                //            }
                //            pathTypes = tempPaths;
                //        }
                //    }
                //    catch { }
                //    if (elementIds == null)
                //    {
                //        elementIds = new List<ElementId>();
                //        paths = new List<string>();
                //        worksheets = new List<string>();
                //        dateTimes = new List<string>();
                //        pathTypes = new List<int>();
                //    }
                //}
            }
            contentOnlyCheckBox.IsChecked = contentOnly;
            contentOnlyCheckBox.ToolTip = "When reloading, only get new content and not style information or modify rows/columns.\nChanging this setting is only remembered while this Manage Excel Links window is open.";
            //linkDataGrid.ItemsSource = LinkedData;
            // Build the table
            BuildTable();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            // Change the ExtensibleStorage if the path was changed
            if (pathChanged)
            {
                paths.Clear();
                pathTypes.Clear();

                foreach (LinkData ld in LinkedData)
                {
                    paths.Add(ld.Path);
                    pathTypes.Add((int)ld.PathType);
                }

                // Get the Schema and entity
                Schema schema = Schema.Lookup(schemaGuid);
                if (schema != null && schema.IsValidObject)
                {
                    //Autodesk.Revit.DB.ExtensibleStorage.Entity entity = null;
                    DataStorage ds = SchemaManager.GetDataStorage(doc);
                    if (ds != null)
                    {
                        ExcelScheduleEntity schedEntity = ds.GetEntity<ExcelScheduleEntity>();
                        if (schedEntity != null)
                        {
                            using (Transaction trans = new Transaction(doc, "Manage Excel Links"))
                            {
                                trans.Start();

                                // Change the Path and PathType Parameters
                                schedEntity.ExcelFilePath = paths.ToList();
                                schedEntity.PathType = pathTypes.ToList();
                                ds.SetEntity(schedEntity);

                                trans.Commit();
                            }
                        }
                    }
                    //try
                    //{
                    //    entity = ds.GetEntity(schema);
                    //}
                    //catch { }

                    //if (entity != null)
                    //{
                    //    Transaction trans = new Transaction(doc, "Manage Excel Links");
                    //    trans.Start();
                    //    entity.Set<IList<string>>("ExcelFilePath", paths);
                    //    entity.Set<IList<int>>("PathType", pathTypes);
                    //    ds.SetEntity(entity);
                    //    trans.Commit();
                    //}
                }
            }

            Close();
        }

        private void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            // Reload the selected schedule
            foreach (var data in linkDataGrid.SelectedItems)
            {
                LinkData selectedRow = (LinkData)data;
                
                for (int i = 0; i < elementIds.Count; i++)
                {
                    try
                    {
                        int intValue = selectedRow.ElementId;
                        if (elementIds[i].IntegerValue == intValue)
                        {
                            // get the full path
                            string docPath;
                            if (doc.IsWorkshared)
                                docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                            else
                                docPath = doc.PathName;
                            string selectedPath = string.Empty;
                            if (selectedRow.PathType == PathType.Absolute)
                                selectedPath = selectedRow.Path;
                            else
                                selectedPath = PathExchange.GetFullPath(selectedRow.Path, docPath);

                            // reload this file
                            if (System.IO.File.Exists(selectedPath))
                            {
                                // read and reload the file.
                                Scheduler creator = new Scheduler();
                                creator.ModifySchedule(doc, elementIds[i], paths[i], worksheets[i], "Reload Excel Schedule", pathTypes[i], contentOnly);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error\n" + ex.Message);
                    }
                }
            }
        }

        private void ReloadFromButton_Click(object sender, RoutedEventArgs e)
        {

            //DataRowView selectedRow = (DataRowView)linkDataGrid.SelectedItems[0];
            LinkData selectedRow = (LinkData)linkDataGrid.SelectedItems[0];
            if (selectedRow != null)
            {
                // Find an Excel File
                System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog()
                {
                    Title = "Reload From an Excel File",
                    Filter = "Excel Files (*.xls; *.xlsx)|*.xls;*.xlsx",
                    RestoreDirectory = true
                };
                

                System.Windows.Forms.DialogResult result = openDlg.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    string excelFilePath = openDlg.FileName;

                    if (System.IO.File.Exists(excelFilePath))
                    {
                        // read the Excel file and create the schedule
                        Excel.Application excelApp = new Excel.Application();
                        Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
                        Excel.Sheets wbWorksheets = workbook.Worksheets;

                        List<WorksheetObject> worksheetObjs = new List<WorksheetObject>();
                        foreach (Excel.Worksheet ws in wbWorksheets)
                        {
                            WorksheetObject wo = new WorksheetObject();
                            string name = ws.Name;
                            wo.Name = name;
                            Excel.Range range = ws.UsedRange;
                            try
                            {
                                range.CopyPicture(Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlBitmap);
                                if (Clipboard.GetDataObject() != null)
                                {
                                    IDataObject data = Clipboard.GetDataObject();
                                    if (data.GetDataPresent(DataFormats.Bitmap))
                                    {
                                        System.Drawing.Image img = (System.Drawing.Image)data.GetData(DataFormats.Bitmap, true);
                                        wo.Image = img;
                                    }
                                }
                            }
                            catch { }
                            worksheetObjs.Add(wo);
                        }

                        // Pop up the worksheet form
                        WorksheetSelectForm wsForm = new WorksheetSelectForm(worksheetObjs, this, doc);
                        wsForm.ShowDialog();

                        if (wsForm.DialogResult.HasValue && wsForm.DialogResult.Value)
                        {
                            for (int i = 0; i < elementIds.Count; i++)
                            {
                                try
                                {
                                    int intValue = selectedRow.ElementId;
                                    if (elementIds[i].IntegerValue == intValue)
                                    {
                                        // read and reload the file.
                                        Scheduler creator = new Scheduler();
                                        creator.ModifySchedule(doc, elementIds[i], excelFilePath, worksheetObj.Name, "Reload Excel Schedule", pathTypes[i], contentOnly);
                                        string docPath;
                                        if (doc.IsWorkshared)
                                            docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                                        else
                                            docPath = doc.PathName;

                                        if ((PathType)pathTypes[i] == PathType.Relative)
                                            paths[i] = PathExchange.GetRelativePath(excelFilePath, docPath);
                                        else
                                            paths[i] = excelFilePath;

                                        worksheets[i] = worksheetObj.Name;
                                        System.IO.FileInfo fi = new System.IO.FileInfo(excelFilePath);
                                        dateTimes[i] = fi.LastWriteTimeUtc.ToString();

                                        // Read the schema information
                                        Autodesk.Revit.DB.ExtensibleStorage.Schema schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(schemaGuid);
                                        if (schema != null)
                                        {
                                            Autodesk.Revit.DB.ExtensibleStorage.Entity entity = null;
                                            DataStorage ds = SchemaManager.GetDataStorage(doc);
                                            try
                                            {
                                                entity = ds.GetEntity(schema);
                                            }
                                            catch { }

                                            if (entity != null)
                                            {
                                                Transaction trans = new Transaction(doc, "Update Excel Document");
                                                trans.Start();
                                                entity.Set<IList<string>>("ExcelFilePath", paths);
                                                entity.Set<IList<string>>("WorksheetName", worksheets);
                                                entity.Set<IList<string>>("DateTime", dateTimes);
                                                entity.Set<IList<int>>("PathType", pathTypes);
                                                ds.SetEntity(entity);
                                                trans.Commit();

                                                BuildTable();
                                            }
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                        try
                        {
                            workbook.Close();
                            Marshal.ReleaseComObject(worksheets);
                            //Marshal.ReleaseComObject(worksheet);
                            Marshal.ReleaseComObject(workbook);
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                        }
                        catch { }
                    }
                }
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            //System.Windows.MessageBox.Show("Selection: " + linkDataGrid.SelectedItems[0].GetType().FullName);
            // Strip this schema object from the entity.
            //DataRowView selectedRow = (DataRowView)linkDataGrid.SelectedItems[0];
            if (linkDataGrid.SelectedItems[0] is LinkData selectedLinkData)
            {
                //string cell = selectedRow.Row.ItemArray[3].ToString();
                    
                // Read the schema information
                Autodesk.Revit.DB.ExtensibleStorage.Schema schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(schemaGuid);
                if (schema != null)
                {
                    Autodesk.Revit.DB.ExtensibleStorage.Entity entity = null;
                    DataStorage ds = SchemaManager.GetDataStorage(doc);
                    try
                    {
                        entity = ds.GetEntity(schema);
                    }
                    catch { }

                    if (entity != null)
                    {
                        elementIds = entity.Get<IList<ElementId>>("ScheduleId");
                        paths = entity.Get<IList<string>>("ExcelFilePath");
                        worksheets = entity.Get<IList<string>>("WorksheetName");
                        dateTimes = entity.Get<IList<string>>("DateTime");
                        pathTypes = entity.Get<IList<int>>("PathType");

                        int index = -1;
                        for (int i = 0; i < elementIds.Count; i++)
                        {
                            try
                            {
                                //int intValue = Convert.ToInt32(cell);
                                if (elementIds[i].IntegerValue == selectedLinkData.ElementId)
                                    index = i;
                            }
                            catch { }
                        }
                            
                        if(index >= 0)
                        {
                            elementIds.RemoveAt(index);
                            paths.RemoveAt(index);
                            worksheets.RemoveAt(index);
                            dateTimes.RemoveAt(index);
                            pathTypes.RemoveAt(index);

                            Transaction trans = new Transaction(doc, "Convert Excel Document to Import");
                            trans.Start();

                            if (elementIds.Count > 0)
                            {
                                entity.Set<IList<ElementId>>("ScheduleId", elementIds);
                                entity.Set<IList<string>>("ExcelFilePath", paths);
                                entity.Set<IList<string>>("WorksheetName", worksheets);
                                entity.Set<IList<string>>("DateTime", dateTimes);
                                entity.Set<IList<int>>("PathType", pathTypes);
                                ds.SetEntity(entity);
                            }
                            else
                            {
                                // Delete the entity data
                                ds.DeleteEntity(schema);
                                doc.Delete(ds.Id);
                            }

                            trans.Commit();

                            BuildTable();
                        }
                    }
                }
            }
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DragMove();
            }
            catch { }
            
        }

        private void BuildTable()
        {
            if (LinkedData == null || LinkedData.Count == 0)
            {
                LinkedData = new List<LinkData>();

                linkDataGrid.CanUserAddRows = false;
                linkDataGrid.CanUserReorderColumns = false;
                linkDataGrid.CanUserSortColumns = false;

                for (int i = 0; i < elementIds.Count; i++)
                {
                    try
                    {
                        LinkData ld = new LinkData();
                        Element schedElem = doc.GetElement(elementIds[i]);
                        try
                        {
                            if (schedElem is ViewSchedule vs)
                            {
                                ld.ScheduleName = vs.Name;
                                ld.WorksheetName = worksheets[i];
                                ld.PathType = (PathType)pathTypes[i];
                                ld.Path = paths[i];
                                ld.ElementId = elementIds[i].IntegerValue;
                                ld.DateTime = dateTimes[i];
                                LinkedData.Add(ld);
                            }
                        }
                        catch { }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("Error-\n" + ex.Message);
                    }
                }
                linkDataGrid.ItemsSource = LinkedData;
                UpdateLayout();
            }
            else
            {
                linkDataGrid.ItemsSource = LinkedData;
                UpdateLayout();
            }
            
        }

        private void PathTypeSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            var comboBox = sender as ComboBox;
            int selectedIndex = comboBox.SelectedIndex;
            PathType pt = PathType.Absolute;
            if (selectedIndex == 1)
                pt = PathType.Relative;
            int currentRowIndex = linkDataGrid.Items.IndexOf(linkDataGrid.CurrentItem);
            
            try
            {
                bool empty = false;
                if (doc.PathName == string.Empty)
                    empty = true;
                if (!empty)
                {
                    
                    if (currentRowIndex >= 0)
                    {
                        
                        LinkData ld = LinkedData[currentRowIndex];
                        string newPath = string.Empty;
                        string docPath;
                        if (doc.IsWorkshared)
                            docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                        else
                            docPath = doc.PathName;
                        switch (pt)
                        {
                            case PathType.Absolute:
                                newPath = PathExchange.GetFullPath(ld.Path, docPath);
                                break;
                            case PathType.Relative:
                                newPath = PathExchange.GetRelativePath(ld.Path, docPath);
                                break;
                        }
                        ld.Path = newPath;
                        ld.PathType = pt;
                        
                        // Rebuild list
                        int listLen = LinkedData.Count;
                        List<LinkData> tempList = LinkedData;
                        LinkedData = new List<LinkData>();
                        for (int i = 0; i < listLen; i++)
                        {
                            if (i == currentRowIndex)
                            {
                                LinkedData.Add(ld);
                            }
                            else
                            {
                                LinkedData.Add(tempList[i]);
                            }
                        }
                        
                        pathChanged = true;
                        //BuildTable();
                        UpdateLayout();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:\n\n" + ex.ToString());
            }
        }

        private void ContentOnlyCheckBox_Click(object sender, RoutedEventArgs e)
        {
            // Change the setting
            if(contentOnlyCheckBox.IsChecked.HasValue)
                contentOnly = contentOnlyCheckBox.IsChecked.Value;
        }
    }
}
