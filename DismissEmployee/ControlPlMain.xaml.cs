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

namespace DismissEmployee
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
        
        // Нажата кнопка увольнения сотрудника
        private void btDismissUser_Click(object sender, RoutedEventArgs e)
        {
            if (!_userIsSelected)
            {
                MessageBox.Show("Не выбран пользователь!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (dateDismiss.SelectedDate == null)
            {
                MessageBox.Show("Не указана дата увольнения!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(_pathToConfig))
            {
                MessageBox.Show("Отсутствует конфигурационный файл плагина!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            string errorMesg = "";
            string dnUserToDismiss = dnUserForChange.Text;
            DateTime? dateToSet = dateDismiss.SelectedDate;
            
            Thread t = new Thread(() =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Clear();
                    _processRun.Add(new ProcessRunData { Text = "---- Увольнение сотрудника " + dnUserToDismiss.Split(',')[0].Substring(3) + " ----"});
                    btSelectUser.IsEnabled = false;
                    btDismissUser.IsEnabled = false;
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
                        itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка чтения конфигурационного файла:\r\n" + errorMesg;
                    }));
                    return;
                }
                if (!configData.ContainsKey("mailserver") || !configData.ContainsKey("templateGroupName"))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Отсутсвуют требуемые параметры в конфигурационном файле!!!";
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Чтение конфигурационного файла плагина";
                }));
                #endregion

                #region Установка срока действия учетной записи
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Установка срока действия учетной записи " + dnUserToDismiss.Split(',')[0].Substring(3) + "..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));

                UserPrincipal userToModify = UserPrincipal.FindByIdentity(_principalContext, IdentityType.DistinguishedName, dnUserToDismiss);
                if (userToModify == null)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Не удалось получить информацию по текущему пользователю в АД!!!";
                    }));
                    return;
                }
                try
                {
                    if (dateToSet != null)
                    {
                        userToModify.AccountExpirationDate = new DateTime(
                            dateToSet.Value.Year,
                            dateToSet.Value.Month,
                            dateToSet.Value.Day,
                            21, 0, 0, 0);
                        userToModify.Save();
                    }
                }
                catch (Exception exc)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка установки срока действия учетной записи:\r\n" + exc.Message;
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Установка срока действия учетной записи " + dnUserToDismiss.Split(',')[0].Substring(3);
                }));
                #endregion

                #region Удаление пользователя из групп
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Удаление пользователя из групп " + configData["templateGroupName"] + "*..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("(&(objectClass=group)(member={0}))", dnUserToDismiss);
                    dirSearcher.PropertiesToLoad.Add("distinguishedName");
                    dirSearcher.PropertiesToLoad.Add("sAMAccountName");
                    dirSearcher.PropertiesToLoad.Add("cn");
                    SearchResultCollection searchResults = dirSearcher.FindAll();
                    foreach (SearchResult sr in searchResults)
                    {
                        if (((string)sr.Properties["sAMAccountName"][0]).Contains(configData["templateGroupName"]) || ((string)sr.Properties["cn"][0]).Contains(configData["templateGroupName"]))
                        {
                            dirSearcher.Filter = string.Format("(&(objectClass=group)(distinguishedName={0}))", (string)sr.Properties["distinguishedName"][0]);
                            dirSearcher.PropertiesToLoad.Add("member");
                            SearchResult searchResult = dirSearcher.FindOne();
                            if (searchResult != null)
                            {
                                DirectoryEntry group = searchResult.GetDirectoryEntry();
                                group.Properties["member"].Remove(dnUserToDismiss);
                                group.CommitChanges();
                            }
                        }
                    }
                }
                catch (Exception exc)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка удаления пользователя из групп:\r\n" + exc.Message;
                    }));
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Удаление пользователя из групп " + configData["templateGroupName"] + "*";
                }));
                #endregion

                #region Отключение ActiveSync и WebApp
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Отключение ActiveSync и WebApp у пользователя " + dnUserToDismiss.Split(',')[0].Substring(3) + "..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                DisableActiveSyncWebAppPS(userToModify.SamAccountName, configData["mailserver"], _login, _password, ref errorMesg);
                if (string.IsNullOrWhiteSpace(errorMesg))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка отключения ActiveSync и WebApp:\r\n" + errorMesg;
                    }));
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/DismissEmployee;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Отключение ActiveSync и WebApp у пользователя " + dnUserToDismiss.Split(',')[0].Substring(3);
                }));
                #endregion

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "---- Завершено ----" });
                    dnUserForChange.Text = "";
                    dnUserForChange.FontWeight = FontWeights.Normal;
                    dnUserForChange.TextDecorations = null;
                    btSelectUser.IsEnabled = true;
                    btDismissUser.IsEnabled = true;
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
        // Отключение ActiveSync и WebApp у пользователя
        private void DisableActiveSyncWebAppPS(string userLogin, string mailServerAdr, string login, SecureString password, ref string errorMsg)
        {
            PSCredential credential = new PSCredential(login, password);
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo((new Uri(mailServerAdr)), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
            Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);

            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();

            command.AddCommand("Set-CASMailbox");
            command.AddParameter("Identity", userLogin);
            command.AddParameter("ActiveSyncEnabled", false);
            command.AddParameter("OWAEnabled", false);

            powershell.Commands = command;
            try
            {
                runspace.Open();
                powershell.Runspace = runspace;
                powershell.Invoke();
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
            }
            finally
            {
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
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
        // В поле выбранного пользователя нажата клавиша
        private void dnUserForChange_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btSelectUser.IsEnabled = false;
                btDismissUser.IsEnabled = false;
                string searchStr = dnUserForChange.Text;
                Thread t = new Thread(() =>
                {
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(!(objectClass=computer))(|(cn=*{0}*)(sn=*{0}*)(givenName=*{0}*)(sAMAccountName=*{0}*)(distinguishedName={0})))", searchStr.Trim());
                    dirSearcher.PropertiesToLoad.Add("distinguishedName");
                    SearchResultCollection searchResults = dirSearcher.FindAll();
                    if (searchResults.Count > 0)
                    {
                        if (searchResults.Count > 1)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                btSelectUser.IsEnabled = true;
                                btDismissUser.IsEnabled = true;
                                btSelectUser_Click("dnUserForChange", new RoutedEventArgs());
                            }));
                        }
                        else
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                dnUserForChange.Text = (string)searchResults[0].Properties["distinguishedName"][0];
                                _userIsSelected = true;
                                dnUserForChange.FontWeight = FontWeights.Bold;
                                dnUserForChange.TextDecorations = TextDecorations.Underline;
                                btSelectUser.IsEnabled = true;
                                btDismissUser.IsEnabled = true;
                            }));
                        }
                    }
                    else
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            btSelectUser.IsEnabled = true;
                            btDismissUser.IsEnabled = true;
                            btSelectUser_Click("dnUserForChange", new RoutedEventArgs());
                        }));
                    }
                });
                t.IsBackground = true;
                t.Start();
            }
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
    }
}
