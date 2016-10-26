using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace GetInfoFromExchange
{
    public class PlugIn : IPlugin
    {
        private string _PluginName = "GetInfoFromExchange";
        private string _DisplayPluginName = "Информация из Exchange";
        private string _PluginDescription = "Данный плагин выбирает информацию по указанному сотруднику из Exchange.";
        private string _Author = "Krivorak V.A.";
        private string _Version = "1.1.0";
        private int _Number = 5;

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
