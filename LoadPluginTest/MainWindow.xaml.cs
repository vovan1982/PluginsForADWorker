using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace LoadPluginTest
{
    public class PluginData
    {
        public int IndexNumber { get; set; }
        public string DisplayName { get; set; }
        public UserControl PLControl { get; set; }
    }

    public class PluginDataComparer : IComparer<PluginData>
    {
        public int Compare(PluginData x, PluginData y)
        {
            if (x == null)
            {
                if (y == null)
                {
                    // If x is null and y is null, they're
                    // equal. 
                    return 0;
                }
                else
                {
                    // If x is null and y is not null, y
                    // is greater. 
                    return -1;
                }
            }
            else
            {
                // If x is not null...
                //
                if (y == null)
                // ...and y is null, x is greater.
                {
                    return 1;
                }
                else
                {
                    // ...and y is not null, compare the 
                    // lengths of the two strings.
                    //
                    if (x.IndexNumber > y.IndexNumber)
                        return 1;
                    else if (x.IndexNumber < y.IndexNumber)
                        return -1;
                    else
                        return 0;
                }
            }

        }
    }
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<PluginData> _plugins;
        private string _login = "";
        private string _pass = "";
        private string _domain = "";
        private DirectoryEntry _connectedSession;
        private PrincipalContext _principalContext;

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                _connectedSession = new DirectoryEntry("LDAP://" + _domain, _domain + @"\" + _login, _pass);
                object obj = _connectedSession.NativeObject;
                _principalContext = new PrincipalContext(ContextType.Domain, _domain, _login, _pass);
            }
            catch (Exception exp)
            {
                startTabLable.Content = exp.Message;
                return;
            }
            _plugins = LoadPlugins(@"D:\Programming\PluginsForADWorker\CreateUsers\bin\Debug", _domain + @"\" + _login, _pass, _connectedSession, _principalContext);
            AddPluginToForm(tabControl,_plugins);
        }
        // Загрузка плагинов из указанной дирректории
        private List<PluginData> LoadPlugins(string path, string login, string pass, DirectoryEntry entry, PrincipalContext context)
        {
            List<PluginData> plugins = new List<PluginData>();
            foreach (var file in Directory.EnumerateFiles(path, "*.dll", SearchOption.AllDirectories))
            {
                try
                {
                    int index = file.ToString().LastIndexOf('\\');
                    string namespase = file.ToString().Substring(index, (file.ToString().Length - index)).TrimEnd(new char[] { '.', 'd', 'l', 'l' }).TrimStart('\\');
                    string configPath = file.ToString().Substring(0, index) + "\\config.ini";
                    if (!File.Exists(configPath))
                        configPath = "";
                    Assembly a = Assembly.LoadFile(file); // dll file
                    if (a == null) continue;
                    Type t = a.GetType(namespase + ".PlugIn"); // namespace - "MyPlayers" , class - "Player"
                    if (t == null) continue;
                    Object instance = Activator.CreateInstance(t);
                    if (instance == null) continue;
                    MethodInfo m = t.GetMethod("PlControl"); // method
                    if (m == null) continue;
                    PropertyInfo f = t.GetProperty("Number");
                    if (f == null) continue;
                    PropertyInfo n = t.GetProperty("DisplayPluginName");
                    if (n == null) continue;
                    string name = (string)n.GetValue(instance, null);
                    int number = (int)f.GetValue(instance, null);
                    UserControl control = (UserControl)m.Invoke(instance, new object[] { login, pass, configPath, entry, context });
                    if (control != null)
                    {
                        plugins.Add(new PluginData { IndexNumber = number, DisplayName = name, PLControl = control });
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
            return plugins;
        }
        // Создание вкладок на форме и добавление загруженых плагинов
        private void AddPluginToForm(TabControl control, List<PluginData> plugins)
        {
            if (plugins.Count > 0)
            {
                plugins.Sort(new PluginDataComparer());
                for (int i = 0; i < _plugins.Count; i++)
                {
                    TabItem tab = new TabItem();
                    tab.Header = plugins[i].DisplayName;
                    tab.Content = plugins[i].PLControl;
                    control.Items.Add(tab);
                }
            }
        }
    }
}
