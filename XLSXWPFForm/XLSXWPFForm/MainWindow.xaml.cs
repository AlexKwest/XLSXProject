using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
        List<OperatorModel> operatorModels;

        public const string poteriashka = "потеряшки";

        public MainWindow()
        {
            InitializeComponent();
        }
        private void submit_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = programDefaultDirectory;
            openFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (openFileDialog.ShowDialog() == true)
            {
                demoPathIn = openFileDialog.FileName;
                btnUploadFirstFile.IsEnabled = true;
            
                logic = new Logic(demoPathIn);
                EnumResult.InputOklad result;

                if (MessageBox.Show("Ввести Оклад по умолчанию \n или ввести вручную?", "Выберите", MessageBoxButton.YesNo,
                    MessageBoxImage.Question) == MessageBoxResult.Yes)
                    result = EnumResult.InputOklad.Default;
                else
                    result = EnumResult.InputOklad.Input;

                operatorModels = logic.SetOperatorList(result);
                lstNameOperator.ItemsSource = operatorModels;

                //.Select(p=>p).ToList<OperatorModel>();
                //this.MyDatagrid.ItemsSource = operatorModels;
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
    }
}
