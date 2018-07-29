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
    /// Логика взаимодействия для FormDialogs.xaml
    /// </summary>
    public partial class FormDialogs : Window
    {
        public static bool Result = false;
        public FormDialogs()
        {
            InitializeComponent();
           // this.Closing += Window_Closing;
        }

        public void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
          
        }

        private void btnYes_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        private void btnNo_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
