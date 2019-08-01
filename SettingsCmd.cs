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