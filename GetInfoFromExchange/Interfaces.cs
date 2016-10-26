using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
namespace GetInfoFromExchange
{
    public interface IPlugin
    {
        string PluginName { get; } // имя плагина
        string DisplayPluginName { get; } // имя плагина, которое отображается
        string PluginDescription { get; } // описание плагина
        string Author { get; } // имя автора
        string Version { get; } // версия
        int Number { get; } // порядковый номер плагина
        System.Windows.Controls.UserControl PlControl(string login, string pass, string pathToConfig, DirectoryEntry entry, PrincipalContext context);
    }
}
