using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace DismissEmployee
{
    public class PlugIn : IPlugin
    {
        private string _PluginName = "DismissEmployee";
        private string _DisplayPluginName = "Уволить сотрудника";
        private string _PluginDescription = "Данный плагин выполняет все действия с АД и Exchange при увольнении сотрудника";
        private string _Author = "Krivorak V.A.";
        private string _Version = "1.2.0";
        private int _Number = 1;

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
