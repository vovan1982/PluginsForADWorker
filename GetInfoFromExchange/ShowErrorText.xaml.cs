using System.Windows;

namespace GetInfoFromExchange
{
    /// <summary>
    /// Логика взаимодействия для ShowErrorText.xaml
    /// </summary>
    public partial class ShowErrorText : Window
    {
        public ShowErrorText(string showText)
        {
            InitializeComponent();
            error.Text = showText;
        }

        private void btOK_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
