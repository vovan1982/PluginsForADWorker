using System.Windows;

namespace DismissEmployee
{
    /// <summary>
    /// Логика взаимодействия для ShowLogText.xaml
    /// </summary>
    public partial class ShowLogText : Window
    {
        public ShowLogText(string showText)
        {
            InitializeComponent();
            log.Text = showText;
        }

        private void btOK_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
