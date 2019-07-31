using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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
