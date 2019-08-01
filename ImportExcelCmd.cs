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
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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
