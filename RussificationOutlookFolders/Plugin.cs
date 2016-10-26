using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace RussificationOutlookFolders
{
    public class PlugIn : IPlugin
    {
        private string _PluginName = "RussificationOutlookFolders";
        private string _DisplayPluginName = "Руссификация папок Outlook";
        private string _PluginDescription = "Данный плагин выполняет руссификацию папок Outlook сотрудника";
        private string _Author = "Krivorak V.A.";
        private string _Version = "1.1.0";
        private int _Number = 6;

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
