using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

using Autodesk.Revit.DB;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.ExtensibleStorage;
using RevitCommon.Attributes;
using RevitCommon.ElementExtensions;


namespace LINE.Revit
{
    public class Scheduler
    {

        private static readonly string dsName = Properties.Settings.Default.DataStorageName;
        static readonly double pointWidthInches = 0.0138888889;
        private static readonly Guid schemaGUID = (typeof(ExcelScheduleEntity).GetCustomAttributes(typeof(SchemaAttribute), true).FirstOrDefault() as SchemaAttribute)?.GUID ?? Guid.Empty;

        Document _doc;
        string docPath = null;
        string excelFilePath = null;
        Excel.Worksheet worksheet = null;
        string workSheetName;
        bool linkFile = true;

        static WorksheetObject selectedWorksheet = null;

        //private bool appOpen = false;
        //private bool workbookOpen = false;

        public WorksheetObject WorksheetObj
        {
            get { return selectedWorksheet; }
            set { selectedWorksheet = value; }
        }

        public bool Link
        {
            get { return linkFile; }
            set { linkFile = value; }
        }

        // Create a new schedule
        public ViewSchedule CreateSchedule(string filePath, UIDocument uidoc)
        {

            ViewSchedule sched = null;
            _doc = uidoc.Document;
            
            if (uidoc.Document.IsWorkshared)
                docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(uidoc.Document.GetWorksharingCentralModelPath());
            else
                docPath = uidoc.Document.PathName;
            
            
            excelFilePath = filePath;
            if (File.Exists(excelFilePath))
            {
                // read the Excel file and create the schedule
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly: true);
                Excel.Sheets worksheets = workbook.Worksheets;
                
                List<WorksheetObject> worksheetObjs = new List<WorksheetObject>();
                foreach (Excel.Worksheet ws in worksheets)
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
                WorksheetSelectForm wsForm = new WorksheetSelectForm(worksheetObjs, this, _doc);


                // Revit version
                int version = Convert.ToInt32(uidoc.Application.Application.VersionNumber);

                // Get the Revit window handle
                IntPtr handle = IntPtr.Zero;
                if (version < 2019)
                    handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                else
                    handle = uidoc.Application.GetType().GetProperty("MainWindowHandle") != null
                        ? (IntPtr)uidoc.Application.GetType().GetProperty("MainWindowHandle").GetValue(uidoc.Application)
                        : IntPtr.Zero;
                System.Windows.Interop.WindowInteropHelper wih = new System.Windows.Interop.WindowInteropHelper(wsForm) { Owner = handle };

                //Show the Worksheet Select form
                wsForm.ShowDialog();
                if (wsForm.DialogResult.HasValue && wsForm.DialogResult.Value)
                {
                    
                    foreach (Excel.Worksheet ws in worksheets)
                    {
                        if (ws.Name == selectedWorksheet.Name)
                        {
                            worksheet = ws;
                            break;
                        }
                    }
                }
                else
                    worksheet = null;

                if (worksheet != null)
                {
                    workSheetName = worksheet.Name;
                    Transaction trans = new Transaction(_doc, "Create Schedule");
                    trans.Start();

                    // Create the schedule
                    sched = ViewSchedule.CreateSchedule(_doc, new ElementId(-1));
                    sched.Name = worksheet.Name;
                    
                    // Add a single parameter for data, Assembly Code
                    ElementId assemblyCodeId = new ElementId(BuiltInParameter.UNIFORMAT_DESCRIPTION);
                    ScheduleFieldId fieldId = null;
                    foreach (SchedulableField sField in sched.Definition.GetSchedulableFields())
                    {
                        ElementId paramId = sField.ParameterId;
                        
                        if (paramId == assemblyCodeId)
                        {
                            ScheduleField field = sched.Definition.AddField(sField);
                            fieldId = field.FieldId;
                            break;
                        }
                        
                    }

                    if (fieldId != null && sched.Definition.GetFieldCount() > 0)
                    {
                        

                        ScheduleDefinition schedDef = sched.Definition;
                        
                        // Add filters to hide all elements in the schedule, ie make sure nothing shows up in the body.
                        ScheduleFilter filter0 = new ScheduleFilter(fieldId, ScheduleFilterType.Equal, "NO VALUES FOUND");
                        ScheduleFilter filter1 = new ScheduleFilter(fieldId, ScheduleFilterType.Equal, "ALL VALUES FOUND");
                        schedDef.AddFilter(filter0);
                        schedDef.AddFilter(filter1);
                        
                        // Turn off the headers
                        schedDef.ShowHeaders = false;
                        
                        // Fill out the schedule from Excel data
                        AddScheduleData(filePath, sched, _doc, PathType.Absolute, false);
                        
                    }
                   


                    if (linkFile)
                        AssignSchemaData(sched.Id, workSheetName, _doc);

                    trans.Commit();
                }

                //workbook.Close();
                workbook.Close(false);
                Marshal.ReleaseComObject(worksheets);
                if(worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            return sched;
        }

        private Excel.Range ActualUsedRange(Excel.Worksheet ws)
        {
            Excel.Range range = null;

            // find the last used row aor column
            int lastColumn = ws.Cells.Find(What: "*", After: ws.get_Range("A1"), SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlPrevious).Column;
            int lastRow = ws.Cells.Find(What: "*", After: ws.get_Range("A1"), SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious).Row;

            try
            {
                Excel.Range start = ws.Cells[1, 1];
                Excel.Range end = ws.Cells[lastRow, lastColumn];
                range = ws.get_Range(start, end);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Test", ex.Message);
            }
            return range;
        }

        private bool AssignSchemaData(ElementId scheduleId, string worksheetName, Document doc)
        {
            // TODO: At some point it may be good to adjust this section. It's currently trying to read the Entity from two different objects
            // TODO: the original ProjectInformation Entity and the new one attached to a DataStorage object which should be more reliable and safe.


            Entity entity = null;
            ExcelScheduleEntity schedEntity = null;
            DataStorage ds = null;

            List<ElementId> elementIds = new List<ElementId>();
            List<string> filePaths = new List<string>();
            List<string> worksheets = new List<string>();
            List<string> dateTimes = new List<string>();
            List<int> pathTypes = new List<int>();


            // Check to see if the ProjectInfo has an Entity stored in it.
            Schema schema = Schema.Lookup(schemaGUID);
            if (null != schema && schema.IsValidObject)
            {
                entity = _doc.ProjectInformation.GetEntity(schema);
                bool purgeFromProjInfo = false;
                if (entity.IsValid())
                {
                    purgeFromProjInfo = true;
                    // Retrieve the data from it
                    elementIds = entity.Get<IList<ElementId>>("ScheduleId").ToList();
                    filePaths = entity.Get<IList<string>>("ExcelFilePath").ToList();
                    worksheets = entity.Get<IList<string>>("WorksheetName").ToList();
                    dateTimes = entity.Get<IList<string>>("DateTime").ToList();
                    pathTypes = entity.Get<IList<int>>("PathType").ToList();

                    // Delete the entity, we should be transitioning over to the DataStorage object.
                    _doc.ProjectInformation.DeleteEntity(schema);
                }

                // See if a datastorage object already exists for this.
                ICollection<ElementId> dsCollector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).ToElementIds();
                foreach (ElementId eid in dsCollector)
                {
                    DataStorage dStor = doc.GetElement(eid) as DataStorage;
                    if (dStor.Name == dsName)
                    {
                        ds = dStor;
                        break;
                    }
                }

                if (purgeFromProjInfo)
                    doc.ProjectInformation.DeleteEntity(schema);

            }

            // Create the dataStorage if necessary
            if (ds == null)
            {
                ds = DataStorage.Create(doc);
                ds.Name = dsName;
            }

            
            schedEntity = ds.GetEntity<ExcelScheduleEntity>();

            if (schedEntity == null)
            {
                // build out the entity and give it default empty lists.
                schedEntity = new ExcelScheduleEntity()
                {
                    ExcelFilePath = filePaths,
                    DateTime = dateTimes,
                    PathType = pathTypes,
                    ScheduleId = elementIds,
                    WorksheetName = worksheets
                };
            }

            schedEntity.ScheduleId.Add(scheduleId);
            schedEntity.ExcelFilePath.Add(excelFilePath);
            schedEntity.WorksheetName.Add(worksheetName);
            FileInfo fi = new FileInfo(excelFilePath);
            schedEntity.DateTime.Add(fi.LastWriteTimeUtc.ToString());
            schedEntity.PathType.Add(0);
            ds.SetEntity(schedEntity);

            return true;
        }

        private void ModifySchemaData(Schema schema, ElementId scheduleId)
        {
            DataStorage ds = SchemaManager.GetDataStorage(_doc);
            if (ds == null || !ds.IsValidObject)// Build data storage if necessary.
                return;

            ExcelScheduleEntity schedEntity = ds.GetEntity<ExcelScheduleEntity>();
            if (schedEntity == null)
                return;

            int index = -1;
            for (int i = 0; i < schedEntity.ScheduleId.Count; i++)
            {
                if (schedEntity.ScheduleId[i].IntegerValue == scheduleId.IntegerValue)
                {
                    index = i;
                    break;
                }
            }

            if (index == -1)
                return;

            schedEntity.DateTime[index] = new FileInfo(schedEntity.ExcelFilePath[index]).LastWriteTimeUtc.ToString();
            ds.SetEntity(schedEntity);
        }

        public void ModifySchedule(Document doc, ElementId scheduleId, string file, string worksheet, string transactionMsg, int pathType, bool contentOnly)
        {
            // Set the Pathtype
            PathType pt = (PathType)pathType;

            // Clear the current schedule of data/formatting and then rebuild it.
            ViewSchedule sched = null;
            _doc = doc;
            try
            {
                Element schedElem = doc.GetElement(scheduleId);
                sched = schedElem as ViewSchedule;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }

            if (sched != null)
            {
                
                Transaction trans = new Transaction(doc, transactionMsg);
                trans.Start();
                if (!contentOnly)
                {
                    // Get the header body to create the necessary rows and columns
                    TableSectionData headerData = sched.GetTableData().GetSectionData(SectionType.Header);
                    int rowCount = headerData.NumberOfRows;
                    int columnCount = headerData.NumberOfColumns;
                    for (int i = 1; i < columnCount; i++)
                    {
                        try
                        {
                            headerData.RemoveColumn(1);
                        }
                        catch { }
                    }
                    for (int i = 1; i < rowCount; i++)
                    {
                        try
                        {
                            headerData.RemoveRow(1);
                        }
                        catch { }
                    }

                    // Make sure the name is up to date
                    sched.Name = worksheet;
                }
                // Add the new schedule data in
                AddScheduleData(file, sched, doc, pt, contentOnly);
                trans.Commit();
            }
            else
            {
                TaskDialog errorDialog = new TaskDialog("Error")
                {
                    MainInstruction = $"Schedule ({worksheet}) could not be found. Remove from update list?",
                    CommonButtons = TaskDialogCommonButtons.No | TaskDialogCommonButtons.Yes,
                    DefaultButton = TaskDialogResult.Yes
                };
                TaskDialogResult result = errorDialog.Show();

                if (result == TaskDialogResult.Yes)
                {
                    try
                    {
                        // Find an Excel File
                        // Get the schema
                        Schema schema = Schema.Lookup(schemaGUID);
                        //Entity entity = null;
                        DataStorage ds = SchemaManager.GetDataStorage(_doc);
                        ExcelScheduleEntity schedEntity = ds.GetEntity<ExcelScheduleEntity>();
                        
                        if (schedEntity != null)
                        {
                            int index = -1;
                            for (int i = 0; i < schedEntity.ScheduleId.Count; i++)
                            {
                                if (schedEntity.ScheduleId[i].IntegerValue == scheduleId.IntegerValue)
                                {
                                    index = i;
                                    break;
                                }
                            }

                            if (index != -1)
                            {
                                using (Transaction trans = new Transaction(doc, "Remove Schedule Excel Link Data"))
                                {
                                    trans.Start();

                                    // Check if there are more linked items than the one we found
                                    if (schedEntity.ScheduleId.Count > 1)
                                    {
                                        // Cull the index
                                        schedEntity.ScheduleId.RemoveAt(index);
                                        schedEntity.DateTime.RemoveAt(index);
                                        schedEntity.ExcelFilePath.RemoveAt(index);
                                        schedEntity.PathType.RemoveAt(index);
                                        schedEntity.WorksheetName.RemoveAt(index);

                                        // Set the entity back to the DS
                                        ds.SetEntity(schedEntity);
                                    }
                                    // If we only have one item and we're removing it, just delete the DataStorage and entity
                                    else
                                    {
                                        ds.DeleteEntity<ExcelScheduleEntity>();
                                        doc.Delete(ds.Id);
                                    }

                                    trans.Commit();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error", ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sched"></param>
        /// <param name="doc"></param>
        /// <param name="pt"></param>
        /// <param name="contentOnly"></param>
        public void AddScheduleData(string filePath, ViewSchedule sched, Document doc, PathType pt, bool contentOnly)
        {
            
            string docPath;
            if (doc.IsWorkshared)
                docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
            else
                docPath = doc.PathName;

            string fullPath;
            if (pt == PathType.Absolute)
                fullPath = filePath;
            else
                fullPath = PathExchange.GetFullPath(filePath, docPath);

            // Get the file path
            excelFilePath = fullPath;
            if (!File.Exists(excelFilePath))
                return;

            
            // read the Excel file and create the schedule
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Sheets worksheets = workbook.Worksheets;
            worksheet = null;
            foreach (Excel.Worksheet ws in worksheets)
            {
                if (ws.Name.Trim() == sched.Name.Trim())
                    worksheet = ws;
            }

            if (worksheet == null)
                return;

            //TaskDialog.Show("Test", "Worksheet found");
            // Find the ThinLine linestyle
            CategoryNameMap lineSubCats = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines).SubCategories;
            ElementId thinLineStyle = new ElementId(-1);
            ElementId hairlineStyle = new ElementId(-1);
            ElementId thinStyle = new ElementId(-1);
            ElementId mediumStyle = new ElementId(-1);
            ElementId thickStyle = new ElementId(-1);
            foreach (Category style in lineSubCats)
            {
                if (style.Name == "Thin Lines")
                {
                    thinLineStyle = style.Id;
                }

                if (style.GetGraphicsStyle(GraphicsStyleType.Projection).Id.IntegerValue == Properties.Settings.Default.hairlineInt)
                {
                    hairlineStyle = style.Id;
                }
                else if (style.GetGraphicsStyle(GraphicsStyleType.Projection).Id.IntegerValue == Properties.Settings.Default.thinInt)
                {
                    thinStyle = style.Id;
                }
                else if (style.GetGraphicsStyle(GraphicsStyleType.Projection).Id.IntegerValue == Properties.Settings.Default.mediumInt)
                {
                    mediumStyle = style.Id;
                }
                else if (style.GetGraphicsStyle(GraphicsStyleType.Projection).Id.IntegerValue == Properties.Settings.Default.thickInt)
                {
                    thickStyle = style.Id;
                }
            }

            if (hairlineStyle.IntegerValue == -1)
                hairlineStyle = thinLineStyle;
            if (thinStyle.IntegerValue == -1)
                thinStyle = thinLineStyle;
            if (mediumStyle.IntegerValue == -1)
                mediumStyle = thinLineStyle;
            if (thickStyle.IntegerValue == -1)
                thickStyle = thinLineStyle;



            // Find out how many rows and columns we need in the schedule
            Excel.Range rng = ActualUsedRange(worksheet);

            Excel.Range range = rng;
            int rowCount = range.Rows.Count;
            int columnCount = range.Columns.Count;

            // Get the schedule body to set the overall width
            TableSectionData bodyData = sched.GetTableData().GetSectionData(SectionType.Body);
            if (!contentOnly)
            {
                double schedWidth = range.Columns.Width;
                try
                {
                    bodyData.SetColumnWidth(0, (schedWidth * pointWidthInches) / 12);
                }
                catch { }
            }

            // Get the header body to create the necessary rows and columns
            TableSectionData headerData = sched.GetTableData().GetSectionData(SectionType.Header);

            if (!contentOnly)
            {
                //TaskDialog.Show("Test: ", "Row Count: " + rowCount.ToString() + "\nColumn Count:  " + columnCount.ToString());
                for (int i = 0; i < columnCount - 1; i++)
                {
                    headerData.InsertColumn(1);
                }
                for (int i = 0; i < rowCount - 1; i++)
                {
                    headerData.InsertRow(1);
                }

                for (int i = 1; i <= headerData.NumberOfColumns; i++)
                {
                    try
                    {
                        Excel.Range cell = worksheet.Cells[1, i];
                        headerData.SetColumnWidth(i - 1, (cell.Width * pointWidthInches) / 12);
                    }
                    catch { }
                }

                for (int i = 1; i <= headerData.NumberOfRows; i++)
                {
                    try
                    {
                        Excel.Range cell = worksheet.Cells[i, 1];

                        headerData.SetRowHeight(i - 1, (cell.Height * pointWidthInches) / 12);
                    }
                    catch { }
                }
            }

            
            
            List<TableMergedCell> mergedCells = new List<TableMergedCell>();
            int errorCount = 0;
            for (int i = 1; i <= headerData.NumberOfRows; i++) // Iterate through rows of worksheet data
            {
                for (int j = 1; j <= headerData.NumberOfColumns; j++) // Iterate through columns of worksheet data
                {
                    // Get the current cell in the worksheet grid
                    Excel.Range cell = worksheet.Cells[i, j];
                    
                    // If adjusting the formatting or adding content is not necessary, 
                    // just update the text content. This is via a UI switch.
                    if (contentOnly)
                    {
                        try
                        {
                            headerData.SetCellText(i-1, j-1, cell.Text);
                            continue;
                        }
                        catch {
                            errorCount++;
                            continue;
                        }
                    }

                    Excel.Font font = cell.Font;
                    Excel.DisplayFormat dispFormat = cell.DisplayFormat;
                    
                    TableCellStyle cellStyle = new TableCellStyle();
                    TableCellStyleOverrideOptions styleOverride = cellStyle.GetCellStyleOverrideOptions();

                    Excel.Border topEdge = cell.Borders.Item[Excel.XlBordersIndex.xlEdgeTop];
                    Excel.Border bottomEdge = cell.Borders.Item[Excel.XlBordersIndex.xlEdgeBottom];
                    Excel.Border leftEdge = cell.Borders.Item[Excel.XlBordersIndex.xlEdgeLeft];
                    Excel.Border rightEdge = cell.Borders.Item[Excel.XlBordersIndex.xlEdgeRight];

                    // Determine Bottom Edge Line Style
                    if (bottomEdge.LineStyle == (int)Excel.XlLineStyle.xlLineStyleNone)
                        cellStyle.BorderBottomLineStyle = new ElementId(-1);
                    else
                    {
                        switch (bottomEdge.Weight)
                        {
                            case (int)Excel.XlBorderWeight.xlHairline:
                                cellStyle.BorderBottomLineStyle = hairlineStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThin:
                                cellStyle.BorderBottomLineStyle = thinStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlMedium:
                                cellStyle.BorderBottomLineStyle = mediumStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThick:
                                cellStyle.BorderBottomLineStyle = thickStyle;
                                break;
                        }
                    }
                    

                    // Determine Top Edge Line Style
                    if (topEdge.LineStyle == (int)Excel.XlLineStyle.xlLineStyleNone)
                        cellStyle.BorderTopLineStyle = new ElementId(-1);
                    else
                    {
                        switch (topEdge.Weight)
                        {
                            case (int)Excel.XlBorderWeight.xlHairline:
                                cellStyle.BorderTopLineStyle = hairlineStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThin:
                                cellStyle.BorderTopLineStyle = thinStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlMedium:
                                cellStyle.BorderTopLineStyle = mediumStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThick:
                                cellStyle.BorderTopLineStyle = thickStyle;
                                break;
                        }
                        
                    }

                    // Determine Left Edge Line Style
                    if (leftEdge.LineStyle == (int)Excel.XlLineStyle.xlLineStyleNone)
                        cellStyle.BorderLeftLineStyle = new ElementId(-1);
                    else
                    {
                        switch (leftEdge.Weight)
                        {
                            case (int)Excel.XlBorderWeight.xlHairline:
                                cellStyle.BorderLeftLineStyle = hairlineStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThin:
                                cellStyle.BorderLeftLineStyle = thinStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlMedium:
                                cellStyle.BorderLeftLineStyle = mediumStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThick:
                                cellStyle.BorderLeftLineStyle = thickStyle;
                                break;
                        }
                    }

                    // Determine Right Edge Line Style
                    if (rightEdge.LineStyle == (int)Excel.XlLineStyle.xlLineStyleNone)
                        cellStyle.BorderRightLineStyle = new ElementId(-1);
                    else
                    {
                        switch (rightEdge.Weight)
                        {
                            case (int)Excel.XlBorderWeight.xlHairline:
                                cellStyle.BorderRightLineStyle = hairlineStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThin:
                                cellStyle.BorderRightLineStyle = thinStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlMedium:
                                cellStyle.BorderRightLineStyle = mediumStyle;
                                break;
                            case (int)Excel.XlBorderWeight.xlThick:
                                cellStyle.BorderRightLineStyle = thickStyle;
                                break;
                        }
                    }
                    // Border Styles are always overridden
                    styleOverride.BorderBottomLineStyle = true;
                    styleOverride.BorderTopLineStyle = true;
                    styleOverride.BorderLeftLineStyle = true;
                    styleOverride.BorderRightLineStyle = true;

                    if (styleOverride.BorderBottomLineStyle || styleOverride.BorderTopLineStyle || 
                       styleOverride.BorderLeftLineStyle || styleOverride.BorderRightLineStyle)
                        styleOverride.BorderLineStyle = true;

                    // Get Background color and font name
                    System.Drawing.Color backGroundColor = System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color);
                    cellStyle.BackgroundColor = new Color(backGroundColor.R, backGroundColor.G, backGroundColor.B);
                    styleOverride.BackgroundColor = true;
                    cellStyle.FontName = cell.Font.Name;
                    styleOverride.Font = true;

                    // Determine Horizontal Alignment
                    // If its not set to left, right or center, do not modify
                    switch (dispFormat.HorizontalAlignment)
                    {
                        case (int)Excel.XlHAlign.xlHAlignLeft:
                            cellStyle.FontHorizontalAlignment = HorizontalAlignmentStyle.Left;
                            styleOverride.HorizontalAlignment = true;
                            break;
                        case (int)Excel.XlHAlign.xlHAlignRight:
                            cellStyle.FontHorizontalAlignment = HorizontalAlignmentStyle.Right;
                            styleOverride.HorizontalAlignment = true;
                            break;
                        case (int)Excel.XlHAlign.xlHAlignGeneral: // No specific style assigned
                            // Check if it's a number which is typically right aligned
                            if (double.TryParse(cell.Text, out double alignTest))
                            {
                                cellStyle.FontHorizontalAlignment = HorizontalAlignmentStyle.Right;
                                styleOverride.HorizontalAlignment = true;
                            }
                            else // Assume text and left align it
                            {
                                cellStyle.FontHorizontalAlignment = HorizontalAlignmentStyle.Left;
                                styleOverride.HorizontalAlignment = true;
                            }
                            break;
                        case (int)Excel.XlHAlign.xlHAlignCenter:
                            cellStyle.FontHorizontalAlignment = HorizontalAlignmentStyle.Center;
                            styleOverride.HorizontalAlignment = true;
                            break;
                    }

                    // Get the vertical alignment of the cell
                    switch (dispFormat.VerticalAlignment)
                    {
                        case (int)Excel.XlVAlign.xlVAlignBottom:
                            cellStyle.FontVerticalAlignment = VerticalAlignmentStyle.Bottom;
                            styleOverride.VerticalAlignment = true;
                            break;
                        case (int)Excel.XlVAlign.xlVAlignTop:
                            cellStyle.FontVerticalAlignment = VerticalAlignmentStyle.Top;
                            styleOverride.VerticalAlignment = true;
                            break;
                        default:
                            cellStyle.FontVerticalAlignment = VerticalAlignmentStyle.Middle;
                            styleOverride.VerticalAlignment = true;
                            break;
                    }

                    switch (dispFormat.Orientation)
                    {
                        case (int)Excel.XlOrientation.xlUpward:
                            cellStyle.TextOrientation = 9;
                            styleOverride.TextOrientation = true;
                            break;
                        case (int)Excel.XlOrientation.xlDownward:
                            cellStyle.TextOrientation = -9;
                            styleOverride.TextOrientation = true;
                            break;
                        case (int)Excel.XlOrientation.xlVertical:
                            cellStyle.TextOrientation = 9;
                            styleOverride.TextOrientation = true;
                            break;
                        default:
                            int rotation = (int) cell.Orientation;
                            if (rotation != (int) Excel.XlOrientation.xlHorizontal)
                            {
                                cellStyle.TextOrientation = rotation;
                                styleOverride.TextOrientation = true;
                            }
                            break;
                    }
                    
                    
                    // Determine Text Size
                    double textSize = Convert.ToDouble(font.Size);
                    //double newTextSize = (textSize / 72) / 12;
                    cellStyle.TextSize = textSize;
                    styleOverride.FontSize = true;

                    // Determine Font Color
                    System.Drawing.Color fontColor = System.Drawing.ColorTranslator.FromOle((int)font.Color);
                    cellStyle.TextColor = new Color(fontColor.R, fontColor.G, fontColor.B);
                    styleOverride.FontColor = true;
                    
                    // NOTES: Bold  is a bool
                    //        Italic is a bool
                    //        Underline is an int
                    cellStyle.IsFontBold = (bool)font.Bold;
                    cellStyle.IsFontItalic = (bool)font.Italic;
                    cellStyle.IsFontUnderline = (int)font.Underline == 2;
                    styleOverride.Bold = true;
                    styleOverride.Italics = true;
                    styleOverride.Underline = true;

                    cellStyle.SetCellStyleOverrideOptions(styleOverride);
                    
                    if (cell.MergeCells == true)
                    {
                        TableMergedCell tmc = new TableMergedCell()
                        {
                            Left = j - 1,
                            Right = cell.MergeArea.Columns.Count - 1,
                            Top = i - 1,
                            Bottom = (i - 1) + cell.MergeArea.Rows.Count - 1
                        };
                    
                        // Check to see if the cell is already merged...
                        bool alreadyMerged = false;
                        foreach (TableMergedCell mergedCell in mergedCells)
                        {
                            bool left = false;
                            bool right = false;
                            bool top = false;
                            bool bottom = false;

                            if (i - 1 >= mergedCell.Top)
                                top = true;
                            if (i - 1 <= mergedCell.Bottom)
                                bottom = true;
                            if (j - 1 >= mergedCell.Left)
                                left = true;
                            if (j - 1 <= mergedCell.Right)
                                right = true;

                            //TaskDialog.Show("MergedCell", string.Format("Top: {0}\nBottom: {1}\nLeft: {2}\nRight: {3}\ni-1: {4}\nj-1: {5}", mergedCell.Top, mergedCell.Bottom, mergedCell.Left, mergedCell.Right, i - 1, j - 1));
                            if (top && bottom && left && right)
                            {
                                alreadyMerged = true;
                                break;
                            }
                        }
                            
                        
                        if (!alreadyMerged)
                        {
                            try
                            {
                                headerData.MergeCells(tmc);
                                headerData.SetCellText(i - 1, j - 1, cell.Text);
                                headerData.SetCellStyle(i - 1, j - 1, cellStyle);
                                j += cell.MergeArea.Columns.Count - 1;
                                mergedCells.Add(tmc);
                            //    TaskDialog.Show("Test", string.Format("This cell [{0},{1}] is merged.\nMerged Area: [{2},{3}]", cell.Row - 1, cell.Column - 1, cell.MergeArea.Rows.Count.ToString(), cell.MergeArea.Columns.Count.ToString()));
                            }
                            catch
                            {
                            }
                        }

                    }
                    else
                    {
                        //TaskDialog.Show("Non Merged", string.Format("This cell is not merged with any others [{0}, {1}]", i - 1, j - 1));
                        try
                        {
                            headerData.SetCellText(i - 1, j - 1, cell.Text);
                            headerData.SetCellStyle(i - 1, j - 1, cellStyle);
                        }
                        catch { }
                    }
                }

            }

            if(errorCount > 0)
                TaskDialog.Show("Warning", "Error reloading content for " + errorCount.ToString() + " cells.\n\nConsider unchecking the \"Content Only\" checkbox and reloading the schedule to force it to rebuild.");

            // Write the Schema to the project
            Schema schema = null;
            try
            {
                schema = Schema.Lookup(schemaGUID);
            }
            catch { }

            ModifySchemaData(schema, sched.Id);
            


            workbook.Close(false);
            Marshal.ReleaseComObject(worksheets);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
    }
}
