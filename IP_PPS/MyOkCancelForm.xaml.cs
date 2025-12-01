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

namespace IP_PPS
{
    /// <summary>
    /// Логика взаимодействия для MyOkCancelForm.xaml
    /// </summary>
    public partial class MyOkCancelForm : Window
    {
        public MyOkCancelForm()
        {
            InitializeComponent();
        }

        public enum Result
        {
            Ok,
            Cancel
        }

        Result result = Result.Cancel;
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            result = Result.Ok;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            result = Result.Cancel;
            this.Close();
        }

        public static Result Show(string title, string message, string okbuttontext, string cancelbuttontext, Result defaultResult = Result.Cancel)
        {
            var form = new MyOkCancelForm();
            form.Title = title;
            form.MessageText.Text = message;
            form.OKButton.Content = okbuttontext;
            form.CancelButton.Content = cancelbuttontext;
            form.ShowDialog();
            return form.result;
        }

    }
}
