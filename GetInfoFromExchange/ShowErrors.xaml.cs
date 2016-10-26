using System.Collections.Generic;
using System.Windows;

namespace GetInfoFromExchange
{
    /// <summary>
    /// Логика взаимодействия для ShowErrors.xaml
    /// </summary>
    public partial class ShowErrors : Window
    {
        public ShowErrors(List<string> errorMessages)
        {
            InitializeComponent();
            error.ItemsSource = errorMessages;
        }

        private void errorItem_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            ShowErrorText _dwSET = new ShowErrorText((string)error.SelectedItem);
            _dwSET.ShowDialog();
        }
    }
}
