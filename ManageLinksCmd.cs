using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace LINE.Revit
{
    [Transaction(TransactionMode.Manual)]
    public class ManageLinksCmd : IExternalCommand
    {
        private readonly Guid schemaGUID = new Guid("91c053bd-edeb-4feb-abac-ef862c311e9d");

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                int version = Convert.ToInt32(commandData.Application.Application.VersionNumber);
                // Construct the form
                ManageExcelLinksForm form = new ManageExcelLinksForm(commandData.Application.ActiveUIDocument.Document, schemaGUID);

                // Get the Revit window handle
                IntPtr handle = IntPtr.Zero;
                if (version < 2019)
                    handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                else
                    handle = commandData.Application.GetType().GetProperty("MainWindowHandle") != null
                        ? (IntPtr)commandData.Application.GetType().GetProperty("MainWindowHandle").GetValue(commandData.Application)
                        : IntPtr.Zero;
                System.Windows.Interop.WindowInteropHelper wih = new System.Windows.Interop.WindowInteropHelper(form) { Owner = handle };

                // Show the form
                form.ShowDialog();

                // Write to home
                RevitCommon.FileUtils.WriteToHome("Excel Import - Manage Links", commandData.Application.Application.VersionName, commandData.Application.Application.Username);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
