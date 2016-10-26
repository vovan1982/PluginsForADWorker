using GetInfoFromExchange.Model;
using System.Windows;
using System.Windows.Media;

namespace GetInfoFromExchange
{
    /// <summary>
    /// Логика взаимодействия для AutoReply.xaml
    /// </summary>
    public partial class AutoReply : Window
    {
        public AutoReply(UserAutoReplayInfo autoReply)
        {
            InitializeComponent();
            if (autoReply != null)
            {
                if (autoReply.State.ToUpper() == "ENABLED")
                {
                    state.Text = "Включено";
                    state.Foreground = new SolidColorBrush(Colors.Green);
                }
                else if (autoReply.State.ToUpper() == "SCHEDULED")
                {
                    state.Text = "Включено на указанный период";
                    state.Foreground = new SolidColorBrush(Colors.Green);
                }
                else
                {
                    state.Text = "Выключено";
                    state.Foreground = new SolidColorBrush(Colors.Red);
                }
                startTime.Text = autoReply.StartDateTime.ToString("dd.MM.yyyy HH:mm:ss");
                endTime.Text = autoReply.EndDateTime.ToString("dd.MM.yyyy HH:mm:ss");
                internalMessageText.Text = autoReply.InternalMessageText;
                externalMessageText.Text = autoReply.ExternalMessageText;
            }
        }
    }
}
