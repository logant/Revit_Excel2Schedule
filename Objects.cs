using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;


namespace LINE.Revit
{
    public class WorksheetObject
    {
        public string Name { get; set; }
        public Image Image { get; set; }
    }

    public enum PathType
    { 
        Absolute = 0, 
        Relative = 1,
    }

    public class LinkData
    {
        public string ScheduleName { get; set; }
        public string WorksheetName { get; set; }
        public PathType PathType { get; set; }
        public string Path { get; set; }
        public int ElementId { get; set; }
        public string DateTime { get; set; }
    }

    public class PathExchange
    {
        public static string GetRelativePath(string fullPath, string docPath)
        {
            string relPath = string.Empty;
            try
            {
                // Get the Excel file and Revit file paths.
                string excelPath = fullPath;
                System.IO.FileInfo revitFile = new System.IO.FileInfo(docPath);
                System.IO.FileInfo excelFile = new System.IO.FileInfo(excelPath);

                System.IO.DirectoryInfo revitDirectory = revitFile.Directory;
                System.IO.DirectoryInfo excelDirectory = excelFile.Directory;

                if (revitDirectory.FullName.ToLower() == excelDirectory.FullName.ToLower())
                {
                    // Set the relative path to just the filename
                    relPath = excelFile.Name;
                }
                else if (excelDirectory.FullName.ToLower().Contains(revitDirectory.FullName.ToLower()))
                {
                    // Relative path is in a subdirectory of the one that contains the revit file.
                    relPath = excelDirectory.FullName.Substring(revitDirectory.FullName.Length + 1) + "\\" + excelFile.Name;
                }
                else
                {
                    
                    string[] revPathArr = docPath.Split(new char[] { '\\' });
                    string[] excelPathArr = excelPath.Split(new char[] { '\\' });

                    bool keepGoing = true;
                    int index = 0;
                    while (keepGoing)
                    {
                        if (revPathArr[index] == excelPathArr[index])
                            index++;
                        else
                            keepGoing = false;
                    }

                    if (index == 0)
                        relPath = excelPath;
                    else
                    {
                        int pathLength = revPathArr.Count() - 1;
                        int retraceLength = pathLength - index;
                        for (int i = 0; i < retraceLength; i++)
                        {
                            relPath += "..\\";
                        }
                        for (int i = retraceLength + 1; i > 1; i--)
                        {
                            relPath += excelPathArr[excelPathArr.Count() - i] + "\\";
                        }

                        relPath += excelPathArr.LastOrDefault();
                    }
                }
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("MESSAGE", ex.ToString());
                Autodesk.Revit.UI.TaskDialog.Show("MESSAGE", "FullPath: " + fullPath + "\nDocPath: " + docPath);
                relPath = fullPath;
            }
            return relPath;
        }

        public static string GetFullPath(string relativePath, string docPath)
        {
            string fullPath = string.Empty;


            System.IO.FileInfo revitFile = new System.IO.FileInfo(docPath);
            System.IO.DirectoryInfo revitDirectory = revitFile.Directory;

            if (System.IO.File.Exists(relativePath))
            {
                fullPath = relativePath;
            }

            else if (!relativePath.Contains("\\"))
            {
                // The file is in the same folder as the revit file
                fullPath = revitDirectory.FullName + "\\" + relativePath;
                //System.Windows.Forms.MessageBox.Show("FullPath: " + fullPath);
            }

            else
            {
                
                string[] relPathArr = relativePath.Split(new char[] { '\\' });
                string[] revPathArr = revitDirectory.FullName.Split(new char[] { '\\' });
                
                if (fullPath == string.Empty)
                {
                    
                    // See if the relative path begins with ..
                    if (relPathArr[0] == "..")
                    {
                        
                        // Count how many retraced paths (..) there are in the relative link
                        int retrace = 0;
                        foreach (string s in relPathArr)
                        {
                            if (s == "..")
                                retrace++;
                        }

                        for (int i = 0; i < revPathArr.Length - retrace; i++)
                        {
                            fullPath += revPathArr[i] + "\\";
                        }

                        for (int i = retrace; i < relPathArr.Length; i++)
                        {
                            fullPath += relPathArr[i] + "\\";
                        }

                        fullPath = fullPath.Substring(0, fullPath.Length - 1);
                    }

                    // Relative path is in a sub folder of the Revit file
                    else
                    {
                        fullPath = revitDirectory.FullName + "\\" + relativePath;
                    }
                }
            }
            
            if (System.IO.File.Exists(fullPath))
                return fullPath;
            else
                return null;
        }
    }

    public class AutoCommitComboBoxColumn : System.Windows.Controls.DataGridComboBoxColumn
    {
        protected override System.Windows.FrameworkElement GenerateEditingElement(System.Windows.Controls.DataGridCell cell, object dataItem)
        {
            var comboBox = (System.Windows.Controls.ComboBox)base.GenerateEditingElement(cell, dataItem);
            comboBox.SelectionChanged += Combobox_SelectionChanged;
            return comboBox;
        }

        public void Combobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            CommitCellEdit((System.Windows.FrameworkElement)sender);
        }
    }
}
