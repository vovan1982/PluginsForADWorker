using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace RussificationOutlookFolders
{
    public class ProcessRunData : INotifyPropertyChanged
    {
        private ImageSource _image;
        private string _text;

        public ImageSource Image
        {
            get { return _image; }
            set 
            { 
                _image = value;
                OnPropertyChanged("Image");
            }
        }
        public string Text
        {
            get { return _text; }
            set 
            { 
                _text = value;
                OnPropertyChanged("Text");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
    /// <summary>
    /// Логика взаимодействия для ControlPlMain.xaml
    /// </summary>
    public partial class ControlPlMain : UserControl
    {
        private string _login; // Логин для соединения с почтовым сервером
        private string _pathToConfig; // Полный путь к конфигурационному файлу плагина
        private SecureString _password; // Пароль для соединения с почтовым сервером
        private DirectoryEntry _ADSession; // Сессия с АД
        private PrincipalContext _principalContext; // Контекст с АД
        private Dictionary<string, string> configData; // Данные конфигурационного файла
        private ObservableCollection<ProcessRunData> _processRun; // Лог
        private bool _userIsSelected; // Признак того что пользователь выбран

        // Конструктор
        public ControlPlMain(IPlugin plug, string login, string pass, string pathToConfig, DirectoryEntry entry, PrincipalContext context)
        {
            InitializeComponent();
            _login = login;
            _password = new SecureString();
            foreach (char x in pass)
            {
                _password.AppendChar(x);
            }
            _ADSession = entry;
            _principalContext = context;
            _pathToConfig = pathToConfig;
            _userIsSelected = false;
            _processRun = new ObservableCollection<ProcessRunData>();
            log.ItemsSource = _processRun;
            lbVersion.Content = lbVersion.Content + " " + plug.Version;
        }
        
        // Нажата кнопка руссификации папок сотрудника
        private void btRussificationFolder_Click(object sender, RoutedEventArgs e)
        {
            if (!_userIsSelected)
            {
                MessageBox.Show("Не выбран пользователь!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(_pathToConfig))
            {
                MessageBox.Show("Отсутствует конфигурационный файл плагина!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            string errorMesg = "";
            string userToRussification = dnUserForChange.Text;
            
            Thread t = new Thread(() =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Clear();
                    _processRun.Add(new ProcessRunData { Text = "---- Руссификация папок пользователя " + userToRussification + " ----" });
                    btSelectUser.IsEnabled = false;
                    btRussificationFolder.IsEnabled = false;
                }));

                #region Чтение конфигурационного файла плагина
                ProcessRunData itemProcess = new ProcessRunData { Text = "Чтение конфигурационного файла плагина..." };
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Чтение конфигурационного файла плагина..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                configData = GetConfigData(_pathToConfig, ref errorMesg);
                if (!string.IsNullOrWhiteSpace(errorMesg))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка чтения конфигурационного файла:\r\n" + errorMesg;
                    }));
                    return;
                }
                if (!configData.ContainsKey("mailserver"))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Отсутсвуют требуемые параметры в конфигурационном файле!!!";
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Чтение конфигурационного файла плагина";
                }));
                #endregion

                //============================================================
                PSCredential credential = new PSCredential(_login, _password);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo((new Uri(configData["mailserver"])), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
                Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);
                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                //============================================================

                #region Подключение к Exchange Server
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Подключение к Exchange Server..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    runspace.Open();
                    powershell.Runspace = runspace;
                }
                catch (Exception exp)
                {
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка подключения к Exchange Server:\r\n" + exp.Message;
                        btSelectUser.IsEnabled = true;
                        btRussificationFolder.IsEnabled = true;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Подключение к Exchange Server";
                }));
                #endregion

                #region Руссификация папок пользователя
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Руссификация папок пользователя " + userToRussification + "..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));

                try
                {
                    command.AddCommand("Set-MailboxRegionalConfiguration");
                    command.AddParameter("Identity", userToRussification);
                    command.AddParameter("Language", "ru-RU");
                    command.AddParameter("LocalizeDefaultFolderName", true);
                    powershell.Commands = command;
                    powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Text = "";
                        }));
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка руссификации папок пользователя:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text += errorMsg;
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            btSelectUser.IsEnabled = true;
                            btRussificationFolder.IsEnabled = true;
                        }));
                        return;
                    }
                }
                catch (Exception exp)
                {
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка руссификации папок пользователя:\r\n" + exp.Message;
                        btSelectUser.IsEnabled = true;
                        btRussificationFolder.IsEnabled = true;
                    }));
                    return;
                }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/RussificationOutlookFolders;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Руссификация папок пользователя " + userToRussification;
                }));

                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
                #endregion

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "---- Процесс завершён! ----" });
                    dnUserForChange.Text = "";
                    _userIsSelected = false;
                    dnUserForChange.FontWeight = FontWeights.Normal;
                    dnUserForChange.TextDecorations = null;
                    btSelectUser.IsEnabled = true;
                    btRussificationFolder.IsEnabled = true;
                }));
            });
            t.IsBackground = true;
            t.Start();
        }
        // Нажата кнопка выбора сотрудника
        private void btSelectUser_Click(object sender, RoutedEventArgs e)
        {
            SelectUser _dwSU = null;
            if (sender is string)
            {
                if ((string)sender == "dnUserForChange")
                {
                    _dwSU = new SelectUser(_ADSession, _principalContext, dnUserForChange.Text);
                    _dwSU.ShowDialog();
                }
            }
            else
            {
                _dwSU = new SelectUser(_ADSession, _principalContext);
                _dwSU.ShowDialog();
            }
            if (_dwSU.DialogResult == true)
            {
                dnUserForChange.Text = _dwSU.SelectedUser;
                _userIsSelected = true;
                dnUserForChange.FontWeight = FontWeights.Bold;
                dnUserForChange.TextDecorations = TextDecorations.Underline;
            }
        }
        // Чтение конфигурационного файла
        private Dictionary<string, string> GetConfigData(string configFilePach, ref string errorStr)
        {
            Dictionary<string, string> configData = new Dictionary<string, string>();
            string configDataFull = "";

            try
            {
                using (StreamReader sr = new StreamReader(configFilePach))
                {
                    configDataFull = sr.ReadToEnd();
                    string[] configArr = configDataFull.Split('\n');
                    for (int i = 0; i < configArr.Length; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(configArr[i]))
                        {
                            string[] conf = configArr[i].Split('=');
                            configData.Add(conf[0], conf[1].TrimEnd('\r'));
                        }
                    }
                }
            }
            catch (Exception e)
            {
                errorStr += "Ошибка чтения файла конфигурации " + configFilePach + ":\r\n" + e.Message;
                configData = null;
            }
            return configData;
        }
        // Двойное нажание на элемент лога
        private void logItem_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            ShowLogText _dwSLT = new ShowLogText(((ProcessRunData)log.SelectedItem).Text);
            _dwSLT.ShowDialog();
        }
        // Изменено значение поля выбранного пользователя
        private void dnUserForChange_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_userIsSelected)
            {
                _userIsSelected = false;
                dnUserForChange.FontWeight = FontWeights.Normal;
                dnUserForChange.TextDecorations = null;
            }
        }
        // В поле выбранного пользователя нажата клавиша
        private void dnUserForChange_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btSelectUser.IsEnabled = false;
                btRussificationFolder.IsEnabled = false;
                dnUserForChange.IsEnabled = false;
                string searchStr = dnUserForChange.Text;
                Thread t = new Thread(() =>
                {
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(!(objectClass=computer))(|(cn=*{0}*)(sn=*{0}*)(givenName=*{0}*)(sAMAccountName=*{0}*)))", searchStr.Trim());
                    dirSearcher.PropertiesToLoad.Add("sAMAccountName");
                    SearchResultCollection searchResults = dirSearcher.FindAll();
                    if (searchResults.Count > 0)
                    {
                        if (searchResults.Count > 1)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                btSelectUser.IsEnabled = true;
                                btRussificationFolder.IsEnabled = true;
                                dnUserForChange.IsEnabled = true;
                                btSelectUser_Click("dnUserForChange", new RoutedEventArgs());
                            }));
                        }
                        else
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                dnUserForChange.Text = (string)searchResults[0].Properties["sAMAccountName"][0];
                                _userIsSelected = true;
                                dnUserForChange.FontWeight = FontWeights.Bold;
                                dnUserForChange.TextDecorations = TextDecorations.Underline;
                                btSelectUser.IsEnabled = true;
                                btRussificationFolder.IsEnabled = true;
                                dnUserForChange.IsEnabled = true;
                            }));
                        }
                    }
                    else
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            btSelectUser.IsEnabled = true;
                            btRussificationFolder.IsEnabled = true;
                            dnUserForChange.IsEnabled = true;
                            btSelectUser_Click("dnUserForChange", new RoutedEventArgs());
                        }));
                    }
                });
                t.IsBackground = true;
                t.Start();
            }
        }
    }
}
