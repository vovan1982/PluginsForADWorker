using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace AutoReplyAndForwarding
{
    public class PlugIn : IPlugin
    {
        private string _PluginName = "AutoReplyAndForwarding";
        private string _DisplayPluginName = "Установка автоответа и пересылки";
        private string _PluginDescription = "Данный плагин выполняет установку автоответа и пересылку писем в Exchange выбраному сотруднику.";
        private string _Author = "Krivorak V.A.";
        private string _Version = "1.3.0";
        private int _Number = 4;

        public System.Windows.Controls.UserControl PlControl(string login, string pass, string pathToConfig, DirectoryEntry entry, PrincipalContext context)
        {
            return new ControlPlMain(this, login, pass, pathToConfig, entry, context);
        }

        public string PluginName
        {
            get
            {
                return _PluginName;
            }
        }

        public string DisplayPluginName
        {
            get
            {
                return _DisplayPluginName;
            }
        }

        public string PluginDescription
        {
            get
            {
                return _PluginDescription;
            }
        }

        public string Author
        {
            get
            {
                return _Author;
            }
        }

        public string Version
        {
            get
            {
                return _Version;
            }
        }

        public int Number
        {
            get
            {
                return _Number;
            }
        }
    }
}
