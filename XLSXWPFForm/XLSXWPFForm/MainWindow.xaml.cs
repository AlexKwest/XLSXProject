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

        static string[] demoPathIn = { System.IO.Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).FullName, "data", "Operators.xlsx") };
        static string demoPathOut = System.IO.Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).FullName, "data", "Result.xlsx");
        public const string poteriashka = "потеряшки";

        OpenFileDialog openFileDialog;
        LogicOperator logicOperator;
        public ObservableCollection<OperatorFirstMonthModel> operatorFirstResult; //модель первого месяца
        public ObservableCollection<OperatorTwoMonthModel> operatorTwoResult;  //модель второго месяца

        public MainWindow()
        {
            InitializeComponent();
        }

        //TODO: Из двух Сейвов сделать один с выбором названия листа
        private void btnSaveFirstFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.InitialDirectory = programDefaultDirectory;
            saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                demoPathOut = saveFileDialog.FileName;

                try
                {
                    logicOperator.PrintResult("1-15 Операторы", EnumResult.PrintFile.FirsMonth, demoPathOut);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnSaveTwoFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = programDefaultDirectory;
            saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                demoPathOut = saveFileDialog.FileName;
                try
                {
                    logicOperator.PrintResult("16-31 Операторы", EnumResult.PrintFile.TwoMonth, demoPathOut);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnLoadBaseOperator_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.InitialDirectory = programDefaultDirectory;
            openFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (openFileDialog.ShowDialog() == true)
            {
                demoPathIn = openFileDialog.FileNames;
                if (demoPathIn.Length == 1)
                {
                    //Когда один файл значит он базовый и нам нужен очет с 1 по 15 чилсо
                    logicOperator = new LogicOperator(demoPathIn[0]);
                    EnumResult.InputOklad result;
                    result = (new FormDialogs().ShowDialog()) == true ? EnumResult.InputOklad.Input : EnumResult.InputOklad.Default;
                    operatorFirstResult = new ObservableCollection<OperatorFirstMonthModel>(logicOperator.SetOperatorFirstMonthList(result));

                    this.DataContext = from t in operatorFirstResult
                                       where (t.Day > 0 && t.Finished > 0) //|| (t.Days31 > 0 && t.Proideno31 > 0)
                                       select t;
                }
                else if (demoPathIn.Length == 2)
                {
                    //Что-то когда 2 файла значит нам нужен отчет с 16 по 31 число 
                    logicOperator = new LogicOperator(demoPathIn);
                    operatorFirstResult = new ObservableCollection<OperatorFirstMonthModel>(logicOperator.GetFirstMonth());
                    operatorTwoResult = new ObservableCollection<OperatorTwoMonthModel>(logicOperator.GetTwoMonth());

                    //TODO: Работает не правильно. Нужно достать Union для двух частей месяца
                    this.DataContext = from t in operatorFirstResult
                                       where (t.Day > 0 && t.Finished > 0) //|| (t.Days31 > 0 && t.Proideno31 > 0)
                                       select t;
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
            OperatorFirstMonthModel asset = new OperatorFirstMonthModel { Name = "ФИО" };
            operatorFirstResult.Add(asset);
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
            DragMove();
        }
        private void ListViewItem_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {

        }
        private void btnLoadTwoOperator_Click(object sender, RoutedEventArgs e)
        {

        }
        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        private void GetUnionDictynary(object sender, RoutedEventArgs e)
        {
            //// UnionDictynary.UnionResult(operatorTwoModels, operatorFirstResult);
            //// var result = UnionDictynary.Result();

            //UnionDictynary.UnionResult(operatorFirstResult, operatorTwoResult);
            //var result = UnionDictynary.Result();

        }

    }
}
