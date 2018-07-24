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
using System.Windows.Shapes;

namespace XLSXWPFForm
{
    /// <summary>
    /// Логика взаимодействия для InputBoxes.xaml
    /// </summary>
    public partial class InputBoxes : Window
    {
        private string name;
        string defaulttext = "Введите ФИО...";//default textbox content
        string errormessage = "Ошибка Сообщения";//error messagebox content
        string errortitle = "Ошибка Заголовка";//error messagebox heading title
        public string OklasResult;
        bool clicked = false;
        Logic logic;

        public InputBoxes(string name, Logic logic)
        {
            InitializeComponent();
            this.name = name;
            lblName.Content = name;
            this.logic = logic;
            this.Closing += Window_Closing;
        }

        public void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!clicked)
            {
                e.Cancel = true;
                logic.inputOklad = EnumResult.InputOklad.Default;
            }
        }

        private void cmbOklad_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            clicked = true;
            if (cmbOklad.SelectedItem == null)
                MessageBox.Show(errormessage, errortitle);
            else
            {
                ComboBox comboBox = (ComboBox)sender;
                ComboBoxItem selectedItem = (ComboBoxItem)comboBox.SelectedItem;
                OklasResult = selectedItem.Content.ToString();
                this.Close();
            }
            clicked = false;

        }
    }
}
