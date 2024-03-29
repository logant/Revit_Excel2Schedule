﻿#region Header
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

using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace LINE.Revit
{
    /// <summary>
    /// Interaction logic for SettingsForm.xaml
    /// </summary>
    public partial class SettingsForm : Window
    {
        List<Autodesk.Revit.DB.Category> _lineStyles;
        Autodesk.Revit.DB.Document _doc;

        Autodesk.Revit.DB.Category hairlineCat;
        Autodesk.Revit.DB.Category thinCat;
        Autodesk.Revit.DB.Category mediumCat;
        Autodesk.Revit.DB.Category thickCat;

        public SettingsForm(List<Autodesk.Revit.DB.Category> lineStyles, Autodesk.Revit.DB.Document doc)
        {
            _lineStyles = lineStyles;
            _doc = doc;
            InitializeComponent();

            // Get current elementIds
            int hairlineInt = Properties.Settings.Default.hairlineInt;
            int thinInt = Properties.Settings.Default.thinInt;
            int mediumInt = Properties.Settings.Default.mediumInt;
            int thickInt = Properties.Settings.Default.thickInt;
            bool reloadValuesOnly = Properties.Settings.Default.reloadValuesOnly;

            // Set the checkbox
            contentOnlyCheckBox.IsEnabled = false;
            contentOnlyCheckBox.Visibility = Visibility.Hidden;
            //contentOnlyCheckBox.IsChecked = reloadValuesOnly;

            // Hairline combobox
            hairLineComboBox.ItemsSource = _lineStyles;
            hairLineComboBox.DisplayMemberPath = "Name";
            try
            {
                if (hairlineInt != -1)
                {
                    Autodesk.Revit.DB.Element hairlineElem = _doc.GetElement(new Autodesk.Revit.DB.ElementId(hairlineInt));
                    
                    for (int i = 0; i < lineStyles.Count; i++)
                    {
                        Autodesk.Revit.DB.Category cat = lineStyles[i];
                        if (cat.Name.Trim() == hairlineElem.Name.Trim())
                        {
                            hairLineComboBox.SelectedIndex = i;
                            break;
                        }
                    }
                }
                else
                {
                    hairLineComboBox.SelectedIndex = 0;
                }
            }
            catch
            {
                hairLineComboBox.SelectedIndex = 0;
            }

            // Thin Combobox
            thinComboBox.ItemsSource = _lineStyles;
            thinComboBox.DisplayMemberPath = "Name";
            try
            {
                if (thinInt != -1)
                {
                    Autodesk.Revit.DB.Element thinElem = _doc.GetElement(new Autodesk.Revit.DB.ElementId(thinInt));
                    for (int i = 0; i < lineStyles.Count; i++)
                    {
                        Autodesk.Revit.DB.Category cat = lineStyles[i];
                        if (cat.Name.Trim() == thinElem.Name.Trim())
                        {
                            thinComboBox.SelectedIndex = i;
                            break;
                        }
                        else
                        {
                            thinComboBox.SelectedIndex = 0;
                        }
                    }
                }
                else
                {
                    thinComboBox.SelectedIndex = 0;
                }
            }
            catch
            {
                thinComboBox.SelectedIndex = 0;
            }

            mediumComboBox.ItemsSource = _lineStyles;
            mediumComboBox.DisplayMemberPath = "Name";
            try
            {
                if (mediumInt != -1)
                {
                    Autodesk.Revit.DB.Element mediumElem = _doc.GetElement(new Autodesk.Revit.DB.ElementId(mediumInt));
                    for (int i = 0; i < lineStyles.Count; i++)
                    {
                        Autodesk.Revit.DB.Category cat = lineStyles[i];
                        if (cat.Name.Trim() == mediumElem.Name.Trim())
                        {
                            mediumComboBox.SelectedIndex = i;
                            break;
                        }
                        else
                        {
                            mediumComboBox.SelectedIndex = 0;
                        }
                    }
                }
                else
                {
                    mediumComboBox.SelectedIndex = 0;
                }
            }
            catch
            {
                mediumComboBox.SelectedIndex = 0;
            }

            thickComboBox.ItemsSource = _lineStyles;
            thickComboBox.DisplayMemberPath = "Name";
            try
            {
                if (thickInt != -1)
                {
                    Autodesk.Revit.DB.Element thickElem = _doc.GetElement(new Autodesk.Revit.DB.ElementId(thickInt));
                    //Autodesk.Revit.UI.TaskDialog.Show("Test", "Thick ElementName: " + thickElem.Name);
                    for (int i = 0; i < lineStyles.Count; i++)
                    {
                        Autodesk.Revit.DB.Category cat = lineStyles[i];
                        if (cat.Name.Trim() == thickElem.Name.Trim())
                        {
                            thickComboBox.SelectedIndex = i;
                            break;
                        }
                        else
                        {
                            thickComboBox.SelectedIndex = 0;
                        }
                    }
                }
                else
                {
                    thickComboBox.SelectedIndex = 0;
                }
            }
            catch
            {
                thickComboBox.SelectedIndex = 0;
            }
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            // Change settings
            Properties.Settings.Default.hairlineInt = hairlineCat.GetGraphicsStyle(Autodesk.Revit.DB.GraphicsStyleType.Projection).Id.IntegerValue;
            Properties.Settings.Default.thinInt = thinCat.GetGraphicsStyle(Autodesk.Revit.DB.GraphicsStyleType.Projection).Id.IntegerValue;
            Properties.Settings.Default.mediumInt = mediumCat.GetGraphicsStyle(Autodesk.Revit.DB.GraphicsStyleType.Projection).Id.IntegerValue;
            Properties.Settings.Default.thickInt = thickCat.GetGraphicsStyle(Autodesk.Revit.DB.GraphicsStyleType.Projection).Id.IntegerValue;
            if (contentOnlyCheckBox.IsChecked.HasValue)
                Properties.Settings.Default.reloadValuesOnly = contentOnlyCheckBox.IsChecked.Value;
            Properties.Settings.Default.Save();
            Close();
        }

        private void hairLineComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedIndex = ((System.Windows.Controls.ComboBox)sender).SelectedIndex;
            Autodesk.Revit.DB.Category selectedStyle = _lineStyles[selectedIndex];
            hairlineCat = selectedStyle;
        }

        private void thinComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedIndex = ((System.Windows.Controls.ComboBox)sender).SelectedIndex;
            Autodesk.Revit.DB.Category selectedStyle = _lineStyles[selectedIndex];
            thinCat = selectedStyle;
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void mediumComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedIndex = ((System.Windows.Controls.ComboBox)sender).SelectedIndex;
            Autodesk.Revit.DB.Category selectedStyle = _lineStyles[selectedIndex];
            mediumCat = selectedStyle;
        }

        private void thickComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedIndex = ((System.Windows.Controls.ComboBox)sender).SelectedIndex;
            Autodesk.Revit.DB.Category selectedStyle = _lineStyles[selectedIndex];
            thickCat = selectedStyle;
        }
    }
}
