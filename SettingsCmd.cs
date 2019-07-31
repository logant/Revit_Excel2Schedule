using System;
using System.Collections.Generic;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace LINE.Revit
{
    [Transaction(TransactionMode.Manual)]
    public class SettingsCmd : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                int version = Convert.ToInt32(commandData.Application.Application.VersionNumber);
                // Collect the categories
                Category lineCat = commandData.Application.ActiveUIDocument.Document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                CategoryNameMap subCats = lineCat.SubCategories;
                List<Category> lineStyles = new List<Category>();
                foreach (Category style in subCats)
                {
                    lineStyles.Add(style);
                }

                // Sort the linestyles
                lineStyles.Sort((x, y) => x.Name.CompareTo(y.Name));

                // Create the form
                SettingsForm form = new SettingsForm(lineStyles, commandData.Application.ActiveUIDocument.Document);

                // Get the Revit window handle
                IntPtr handle = IntPtr.Zero;
                if (version < 2019)
                    handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                else
                    handle = commandData.Application.GetType().GetProperty("MainWindowHandle") != null
                        ? (IntPtr) commandData.Application.GetType().GetProperty("MainWindowHandle").GetValue(commandData.Application)
                        : IntPtr.Zero;
                System.Windows.Interop.WindowInteropHelper wih = new System.Windows.Interop.WindowInteropHelper(form) {Owner = handle};

                // Show the form
                form.ShowDialog();

                // Write to home
                RevitCommon.FileUtils.WriteToHome("Excel Import - Settings", commandData.Application.Application.VersionName, commandData.Application.Application.Username);

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
    }
}