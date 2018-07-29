using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
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
using XLSXWPFForm.model;

namespace XLSXWPFForm
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static string programDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        static string programDefaultDirectory = System.IO.Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).FullName, "data");

        static string demoPathIn = System.IO.Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).FullName, "data", "Operators.xlsx");
        static string demoPathOut = System.IO.Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).FullName, "data", "Result.xlsx");

        OpenFileDialog openFileDialog;
        Logic logic;
        public ObservableCollection<OperatorModel> operatorModels;
        public OperatorModel operatorResult;

        public const string poteriashka = "потеряшки";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void submit_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = programDefaultDirectory;
            openFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (openFileDialog.ShowDialog() == true)
            {
                demoPathIn = openFileDialog.FileName;
                //btnUploadFirstFile.IsEnabled = true;
            
                logic = new Logic(demoPathIn);
                EnumResult.InputOklad result;

                result = (new FormDialogs().ShowDialog()) == true? EnumResult.InputOklad.Input : EnumResult.InputOklad.Default;
                operatorModels = new ObservableCollection<OperatorModel>(logic.SetOperatorList(result));
               
               //lstOperator.ItemsSource = operatorModels;
               this.DataContext = from t in operatorModels
                                  where t.Days15 > 0 || t.Days31 > 0 
                                  select t;
            }
        }

        private void btnUploadFirstFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.InitialDirectory = programDefaultDirectory;
            saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                demoPathOut = saveFileDialog.FileName;

                try
                {
                    logic.PrintResult("1-15 Операторы", EnumResult.PrintFile.FirsMonth, demoPathOut);
                    logic.PrintResult("16-31 Операторы", EnumResult.PrintFile.TwoMonth, demoPathOut);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void SwitchShowOperators(object sender, RoutedEventArgs e)
        {
            lstOperator.Visibility = lstOperator.Visibility == Visibility.Collapsed ? Visibility.Visible : Visibility.Collapsed;
        }

        private void list_Selected(object sender, RoutedEventArgs e)
        {
            //lstViewOperator.Items.Clear();
            //operatorResult = (OperatorModel)lstNameOperator.SelectedItem;
            //lstViewOperator.Items.Add(operatorResult);
        }
    

    private void OnColumnHeaderClick(object sender, RoutedEventArgs e)
        {
            GridViewColumn column = ((GridViewColumnHeader)e.OriginalSource).Column;
            //piePlotter.PlottedProperty = column.Header.ToString();
        }

        private void AddNewItem(object sender, RoutedEventArgs e)
        {
            OperatorModel asset = new OperatorModel { Name = "ФИО" };
            operatorModels.Add(asset);
        }

        private void btnOpenMenu_Click(object sender, RoutedEventArgs e)
        {
            btnOpenMenu.Visibility = Visibility.Collapsed;
            btnCloseMenu.Visibility = Visibility.Visible;
        }

        private void btnCloseMenu_Click(object sender, RoutedEventArgs e)
        {
            btnOpenMenu.Visibility = Visibility.Visible;
            btnCloseMenu.Visibility = Visibility.Collapsed;
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void ListViewItem_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {

        }
    }
}
