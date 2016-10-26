using CreateUsers.Model;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;

namespace CreateUsers.DialogWindows
{
    /// <summary>
    /// Логика взаимодействия для Report.xaml
    /// </summary>
    public partial class Report : Window
    {
        public Report(ObservableCollection<ProcessRunData> processRun)
        {
            InitializeComponent();
            report.ItemsSource = processRun;
        }

        // Двойное нажание на элемент лога
        private void reportItem_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            ShowLogText _dwSLT = new ShowLogText(((ProcessRunData)report.SelectedItem).Text);
            _dwSLT.ShowDialog();
        }
        // Нажата кнопка в основном окне
        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
    }
}
