//using System.Collections.Generic;
//using System.Collections.ObjectModel;
//using System.Windows;
//using System.Windows.Input;
//using System.Windows.Media;

//namespace LINE.Revit
//{
//    /// <summary>
//    /// Interaction logic for ReloadForm.xaml
//    /// </summary>
//    public partial class ReloadForm : Window
//    {
//        ObservableCollection<WorksheetUpdaterClass> ListContent;

//        public ReloadForm(IList<string> worksheets, List<int> indeces)
//        {
//            InitializeComponent();
//            ListContent = new ObservableCollection<WorksheetUpdaterClass>();
//            int counter = 0;
//            foreach(string s in worksheets)
//            {
//                WorksheetUpdaterClass w = new WorksheetUpdaterClass();
//                w.ScheduleName = s;
//                w.ContentAndFormatting = true;
//                w.ContentOnly = false;
//                w.Index = counter;
//                ListContent.Add(w);
//                counter++;
//            }

//            scheduleListView.ItemsSource = ListContent;
//        }

//        private void okButton_Click(object sender, RoutedEventArgs e)
//        {
//            Close();
//        }


//        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
//        {
//            DragMove();
//        }
//    }

//    public class WorksheetUpdaterClass
//    {
//        public string ScheduleName { get; set; }
//        public int Index { get; set; }

//        private bool contentOnly = true;

//        public bool ContentOnly
//        {
//            get { return contentOnly; }
//            set
//            {
//                contentOnly = value;
//                ContentAndFormatting = !value;
//            }
//        }

//        public bool ContentAndFormatting { get; set; }
//    }
//}
