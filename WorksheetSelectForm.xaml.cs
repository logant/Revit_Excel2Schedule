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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace LINE.Revit
{
    /// <summary>
    /// Interaction logic for WorksheetSelectForm.xaml
    /// </summary>
    public partial class WorksheetSelectForm : Window
    {
        WorksheetObject selectedWorksheet = null;
        Scheduler _parent;
        ManageExcelLinksForm formParent;
        List<WorksheetObject> _objs;
        Autodesk.Revit.DB.Document _doc;

        //LinearGradientBrush brush = null;

        public WorksheetSelectForm(List<WorksheetObject> objs, Scheduler parent, Autodesk.Revit.DB.Document doc)
        {
            _objs = objs;
            _parent = parent;
            _doc = doc;
            InitializeComponent();

            int counter = 0;
            foreach (WorksheetObject wo in _objs)
            {
                if (wo.Image == null)
                    counter++;
            }

            if(counter == _objs.Count)
            {
                this.Height = 160;
                previewImage.Visibility = System.Windows.Visibility.Hidden;
            }
            linkCheckBox.IsChecked = true;
            wsComboBox.ItemsSource = _objs;
            wsComboBox.DisplayMemberPath = "Name";
            wsComboBox.SelectedIndex = 0;
        }

        public WorksheetSelectForm(List<WorksheetObject> objs, ManageExcelLinksForm parent, Autodesk.Revit.DB.Document doc)
        {
            _objs = objs;
            formParent = parent;
            _doc = doc;
            InitializeComponent();

            int counter = 0;
            foreach (WorksheetObject wo in _objs)
            {
                if (wo.Image == null)
                    counter++;
            }

            if (counter == _objs.Count)
            {
                this.Height = 160;
                previewImage.Visibility = System.Windows.Visibility.Hidden;
            }
            linkCheckBox.IsChecked = true;
            linkCheckBox.IsEnabled = false;
            wsComboBox.ItemsSource = _objs;
            wsComboBox.DisplayMemberPath = "Name";
            wsComboBox.SelectedIndex = 0;
        }


        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            try
            {
                _parent.WorksheetObj = null;
            }
            catch { }
            Close();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            if (_parent != null)
            {
                _parent.WorksheetObj = selectedWorksheet;
                _parent.Link = (bool)linkCheckBox.IsChecked;
            }
            else if (formParent != null)
            {
                formParent.Worksheet = selectedWorksheet;
            }

            Close();
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void WsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedIndex = ((System.Windows.Controls.ComboBox)sender).SelectedIndex;
            WorksheetObject wsObj = _objs[selectedIndex];
            selectedWorksheet = wsObj;
            if (wsObj.Image != null)
            {
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(wsObj.Image);
                BitmapSource source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(bitmap.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                previewImage.Source = source;
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            // Collect the categories
            Autodesk.Revit.DB.Category lineCat = _doc.Settings.Categories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_Lines);
            Autodesk.Revit.DB.CategoryNameMap subCats = lineCat.SubCategories;
            List<Autodesk.Revit.DB.Category> lineStyles = new List<Autodesk.Revit.DB.Category>();
            foreach (Autodesk.Revit.DB.Category style in subCats)
            {
                lineStyles.Add(style);
            }

            // Sort the linestyles
            lineStyles.Sort((x, y) => x.Name.CompareTo(y.Name));

            SettingsForm form = new SettingsForm(lineStyles, _doc);
            form.ShowDialog();
        }
    }
}
