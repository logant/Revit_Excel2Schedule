using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace LINE.Revit
{
    [Transaction(TransactionMode.Manual)]
    public class ImportExcelCmd : IExternalCommand
    {
        string excelFilePath = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Find an Excel File
                System.Windows.Forms.OpenFileDialog openDlg = new System.Windows.Forms.OpenFileDialog()
                {
                    Title = "Import an Excel File",
                    Filter = "Excel Files (*.xls; *.xlsx)|*.xls;*.xlsx",
                    RestoreDirectory = true
                };
                
                System.Windows.Forms.DialogResult result = openDlg.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    if (openDlg.FileName != null)
                    {
                        excelFilePath = openDlg.FileName;
                        Scheduler scheduler = new Scheduler();
                        ViewSchedule vs = scheduler.CreateSchedule(excelFilePath, commandData.Application.ActiveUIDocument);
                    }
                }
                else
                {
                    return Result.Cancelled;
                }

                // Write to home
                RevitCommon.FileUtils.WriteToHome("Excel Import", commandData.Application.Application.VersionName, commandData.Application.Application.Username);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
        }
    }
}
