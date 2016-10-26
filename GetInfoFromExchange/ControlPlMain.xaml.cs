using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Security;
using System.Windows;
using System.Windows.Controls;
using GetInfoFromExchange.Model;
using System.Windows.Media;
using System.Threading;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Globalization;
using System.Windows.Input;

namespace GetInfoFromExchange
{
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
        private List<string> _errorMessages; // Список сообщений об ошибках при загрузке информации из Exchange
        private UserExchangeInfo loadedData; // Загруженная информация
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
            _errorMessages = new List<string>();
            loadedData = new UserExchangeInfo();
            lbVersion.Content = lbVersion.Content + " " + plug.Version;
        }

        // Нажата кнопка выбора сотрудника
        private void btSelectUser_Click(object sender, RoutedEventArgs e)
        {
            SelectUser _dwSU = null;
            if (sender is string)
            {
                if ((string)sender == "loginUser")
                {
                    _dwSU = new SelectUser(_ADSession, _principalContext, loginUser.Text);
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
                loginUser.Text = _dwSU.SelectedUser.Login;
                _userIsSelected = true;
                loginUser.FontWeight = FontWeights.Bold;
                loginUser.TextDecorations = TextDecorations.Underline;
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
        // Деактивация формы
        private void isEnableForm(bool enable)
        {
            btSelectUser.IsEnabled = enable;
            loginUser.IsEnabled = enable;
            if (enable && _errorMessages.Count > 0){
                btShowErrors.IsEnabled = true;
                btInboxMessageRule.IsEnabled = false;
                btAutoReplyInfo.IsEnabled = false;
                btMobileDeviceInfo.IsEnabled = false;
            }
            else if (enable && _errorMessages.Count <= 0){
                btShowErrors.IsEnabled = false;
                btInboxMessageRule.IsEnabled = true;
                btAutoReplyInfo.IsEnabled = true;
                btMobileDeviceInfo.IsEnabled = true;
            }
            else
            {
                btShowErrors.IsEnabled = enable;
                btInboxMessageRule.IsEnabled = enable;
                btAutoReplyInfo.IsEnabled = enable;
                btMobileDeviceInfo.IsEnabled = enable;
            }
            btLoadInfo.IsEnabled = enable;
        }
        // Установка загруженных данных на форму
        private void setLoadedData(UserExchangeInfo data)
        {
            titleText.Content = "Информация по сотруднику " + loadedData.User;

            if (data.IsEnableActiveSync)
            {
                activeSyncState.Foreground = new SolidColorBrush(Colors.Green);
                activeSyncState.Text = "Включен";
            }
            else
            {
                activeSyncState.Foreground = new SolidColorBrush(Colors.Red);
                activeSyncState.Text = "Выключен";
            }
            if (data.IsEnableWebApp)
            {
                webAppState.Foreground = new SolidColorBrush(Colors.Green);
                webAppState.Text = "Включен";
            }
            else
            {
                webAppState.Foreground = new SolidColorBrush(Colors.Red);
                webAppState.Text = "Выключен";
            }
            activeSyncPolicy.Text = data.PolicyActiveSync;
            webAppPolicy.Text = data.PolicyWebApp;
            if (data.IsEnableMAPI)
            {
                mapiState.Foreground = new SolidColorBrush(Colors.Green);
                mapiState.Text = "Включен";
            }
            else
            {
                mapiState.Foreground = new SolidColorBrush(Colors.Red);
                mapiState.Text = "Выключен";
            }
            if (data.IsEnablePOP)
            {
                popState.Foreground = new SolidColorBrush(Colors.Green);
                popState.Text = "Включен";
            }
            else
            {
                popState.Foreground = new SolidColorBrush(Colors.Red);
                popState.Text = "Выключен";
            }
            if (data.IsEnableIMAP)
            {
                imapState.Foreground = new SolidColorBrush(Colors.Green);
                imapState.Text = "Включен";
            }
            else
            {
                imapState.Foreground = new SolidColorBrush(Colors.Red);
                imapState.Text = "Выключен";
            }
            forwardingAddress.Text = data.ForwardingAddress;
            forwardingSmtpAddress.Text = data.ForwardingSmtpAddress;
            emailAddress.Text = data.EmailAddress;
            btInboxMessageRule.IsEnabled = true;
            btAutoReplyInfo.IsEnabled = true;
            btMobileDeviceInfo.IsEnabled = true;
        }
        // Нажата кнопка загрузки информации 
        private void btLoadInfo_Click(object sender, RoutedEventArgs e)
        {
            #region Проверка и инициализация входных данных
            if (!_userIsSelected)
            {
                MessageBox.Show("Не выбран пользователь!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            loadedData.User = loginUser.Text;
            string errorMesg = "";
            #endregion

            Thread t = new Thread(() =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    isEnableForm(false);
                }));

                #region Чтение конфигурационного файла плагина
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "Чтение конфигурационного файла плагина...";
                }));
                configData = GetConfigData(_pathToConfig, ref errorMesg);
                if (!string.IsNullOrWhiteSpace(errorMesg))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _errorMessages.Add("Ошибка чтения конфигурационного файла:\r\n" + errorMesg);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                if (!configData.ContainsKey("mailserver"))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _errorMessages.Add("Отсутсвуют требуемые параметры в конфигурационном файле!!!");
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
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
                    statusBarText.Content = "Подключение к Exchange Server...";
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
                        _errorMessages.Add("Ошибка подключения к Exchange Server:\r\n" + exp.Message);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
                }));
                #endregion

                #region Загрузка функций почтового ящика
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "Загрузка функций почтового ящика...";
                }));
                try
                {
                    string[] loadParam = new string[] {"ActiveSyncEnabled", "ActiveSyncMailboxPolicy", "OWAEnabled", "OwaMailboxPolicy", "MAPIEnabled", "PopEnabled", "ImapEnabled" };
                    command.AddCommand("Get-CASMailbox");
                    command.AddParameter("Identity", loadedData.User);
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", loadParam);

                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки функций почтового ящика:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                _errorMessages.Add(errorMsg);
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                            isEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        var cmd = result[0];
                        if (cmd.Properties["ActiveSyncEnabled"].Value.ToString().ToUpper() == "TRUE")
                            loadedData.IsEnableActiveSync = true;
                        else
                            loadedData.IsEnableActiveSync = false;
                        if (cmd.Properties["ActiveSyncMailboxPolicy"].Value != null)
                            loadedData.PolicyActiveSync = cmd.Properties["ActiveSyncMailboxPolicy"].Value.ToString();
                        else
                            loadedData.PolicyActiveSync = "";
                        if (cmd.Properties["OWAEnabled"].Value.ToString().ToUpper() == "TRUE")
                            loadedData.IsEnableWebApp = true;
                        else
                            loadedData.IsEnableWebApp = false;
                        if (cmd.Properties["OwaMailboxPolicy"].Value != null)
                            loadedData.PolicyWebApp = cmd.Properties["OwaMailboxPolicy"].Value.ToString();
                        else
                            loadedData.PolicyWebApp = "";
                        if (cmd.Properties["MAPIEnabled"].Value.ToString().ToUpper() == "TRUE")
                            loadedData.IsEnableMAPI = true;
                        else
                            loadedData.IsEnableMAPI = false;
                        if (cmd.Properties["PopEnabled"].Value.ToString().ToUpper() == "TRUE")
                            loadedData.IsEnablePOP = true;
                        else
                            loadedData.IsEnablePOP = false;
                        if (cmd.Properties["ImapEnabled"].Value.ToString().ToUpper() == "TRUE")
                            loadedData.IsEnableIMAP = true;
                        else
                            loadedData.IsEnableIMAP = false;
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
                        _errorMessages.Add("Ошибка загрузки функций почтового ящика:\r\n" + exp.Message);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
                }));
                #endregion

                #region Загрузка переадресации
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "Загрузка переадресации...";
                }));
                try
                {
                    string[] loadParam = new string[] { "ForwardingAddress", "ForwardingSmtpAddress" };
                    command.AddCommand("Get-Mailbox");
                    command.AddParameter("Identity", loadedData.User);
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", loadParam);

                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки переадресации:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                _errorMessages.Add(errorMsg);
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                            isEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        var cmd = result[0];

                        if (cmd.Properties["ForwardingAddress"].Value != null)
                            loadedData.ForwardingAddress = cmd.Properties["ForwardingAddress"].Value.ToString();
                        else
                            loadedData.ForwardingAddress = "";

                        if (cmd.Properties["ForwardingSmtpAddress"].Value != null)
                            loadedData.ForwardingSmtpAddress = cmd.Properties["ForwardingSmtpAddress"].Value.ToString();
                        else
                            loadedData.ForwardingSmtpAddress = "";
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
                        _errorMessages.Add("Ошибка загрузки переадресации:\r\n" + exp.Message);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
                }));
                #endregion

                #region Загрузка электронных адресов
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "Загрузка электронных адресов...";
                }));
                try
                {
                    command.AddCommand("Get-Mailbox");
                    command.AddParameter("Identity", loadedData.User);
                    command.AddCommand("Select-Object");
                    command.AddParameter("ExpandProperty", "EmailAddresses");
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки электронных адресов:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                _errorMessages.Add(errorMsg);
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                            isEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        loadedData.EmailAddress = "";
                        for (int i = 0; i < result.Count; i++)
                        {
                            loadedData.EmailAddress += result[i].Properties["ProxyAddressString"].Value.ToString() + "\r\n";
                            if (result[i].Properties["ProxyAddressString"].Value.ToString().StartsWith("SMTP"))
                            {
                                loadedData.PrimaryEmailAddres = result[i].Properties["ProxyAddressString"].Value.ToString().Split(':')[1];
                            }
                        }
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
                        _errorMessages.Add("Ошибка загрузки электронных адресов:\r\n" + exp.Message);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
                }));
                #endregion

                #region Загрузка правил входящих сообщений
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "Загрузка правил входящих сообщений...";
                }));
                try
                {
                    string[] loadParam = new string[] { "Enabled", "Priority", "Name", "Description" };
                    loadedData.IndoxMessageRule = new List<InboxMessageRuleInfo>();
                    command.AddCommand("Get-InboxRule");
                    command.AddParameter("Mailbox", loadedData.PrimaryEmailAddres);
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", loadParam);
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки правил входящих сообщений:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                _errorMessages.Add(errorMsg);
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                            isEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        for (int i = 0; i < result.Count; i++)
                        {
                            var cmd = result[i];
                            loadedData.IndoxMessageRule.Add(new InboxMessageRuleInfo { 
                                IsEnabled = cmd.Properties["Enabled"].Value.ToString().ToUpper() == "TRUE" ? true : false,
                                Priority = Int32.Parse(cmd.Properties["Priority"].Value.ToString()),
                                Name = cmd.Properties["Name"].Value.ToString(),
                                Description = cmd.Properties["Description"].Value.ToString()
                            });
                        }
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
                        _errorMessages.Add("Ошибка загрузки правил входящих сообщений:\r\n" + exp.Message);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
                }));
                #endregion

                #region Загрузка информации по установленному автоответу
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "Загрузка информации по установленному автоответу...";
                }));
                try
                {
                    string[] loadParam = new string[] { "AutoReplyState", "StartTime", "EndTime", "InternalMessage", "ExternalMessage" };
                    loadedData.AutoReplay = new UserAutoReplayInfo();
                    command.AddCommand("Get-MailboxAutoReplyConfiguration");
                    command.AddParameter("Identity", loadedData.User);
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", loadParam);
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки информации по установленному автоответу:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                _errorMessages.Add(errorMsg);
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                            isEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        loadedData.AutoReplay = new UserAutoReplayInfo() { 
                            State = result[0].Properties["AutoReplyState"].Value.ToString(),
                            StartDateTime = DateTime.ParseExact(result[0].Properties["StartTime"].Value.ToString(), "dd.MM.yyyy H:mm:ss", CultureInfo.InvariantCulture),
                            EndDateTime = DateTime.ParseExact(result[0].Properties["EndTime"].Value.ToString(), "dd.MM.yyyy H:mm:ss", CultureInfo.InvariantCulture),
                            InternalMessageText = result[0].Properties["InternalMessage"].Value != null ? result[0].Properties["InternalMessage"].Value.ToString() : "",
                            ExternalMessageText = result[0].Properties["ExternalMessage"].Value != null ? result[0].Properties["ExternalMessage"].Value.ToString() : ""
                        };
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
                        _errorMessages.Add("Ошибка загрузки информации по установленному автоответу:\r\n" + exp.Message);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
                }));
                #endregion

                #region Загрузка информации по мобильным устройствам
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "Загрузка информации по мобильным устройствам...";
                }));
                try
                {
                    string[] loadParam = new string[] { "DeviceType", "DeviceFriendlyName", "DeviceModel", "DeviceUserAgent", "LastSyncAttemptTime", "LastSuccessSync" };
                    loadedData.MobileDevice = new List<MobileDeviceInfo>();
                    command.AddCommand("Get-ActiveSyncDeviceStatistics");
                    command.AddParameter("Mailbox", loadedData.PrimaryEmailAddres);
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", loadParam);
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки информации по мобильным устройствам:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                _errorMessages.Add(errorMsg);
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                            isEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        for (int i = 0; i < result.Count; i++)
                        {
                            var cmd = result[i];
                            loadedData.MobileDevice.Add(new MobileDeviceInfo
                            {
                                DeviceType = cmd.Properties["DeviceType"].Value != null ? cmd.Properties["DeviceType"].Value.ToString() : "",
                                DeviceName = cmd.Properties["DeviceFriendlyName"].Value != null ? cmd.Properties["DeviceFriendlyName"].Value.ToString() : "",
                                DeviceModel = cmd.Properties["DeviceModel"].Value != null ? cmd.Properties["DeviceModel"].Value.ToString() : "",
                                DeviceUserAgent = cmd.Properties["DeviceUserAgent"].Value != null ? cmd.Properties["DeviceUserAgent"].Value.ToString() : "",
                                LastSuccessSyncTime = cmd.Properties["LastSyncAttemptTime"].Value != null ? DateTime.ParseExact(cmd.Properties["LastSyncAttemptTime"].Value.ToString(), "dd.MM.yyyy H:mm:ss", CultureInfo.InvariantCulture) : (DateTime?)null,
                                LastSyncTime = cmd.Properties["LastSuccessSync"].Value != null ? DateTime.ParseExact(cmd.Properties["LastSuccessSync"].Value.ToString(), "dd.MM.yyyy H:mm:ss", CultureInfo.InvariantCulture) : (DateTime?)null
                            });
                        }
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
                        _errorMessages.Add("Ошибка информации по мобильным устройствам:\r\n" + exp.Message);
                        statusBarText.Content = "В процессе загрузки возникли ошибки!!!";
                        isEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusBarText.Content = "";
                }));
                #endregion

                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
                command.Clear();
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    setLoadedData(loadedData);
                    isEnableForm(true);
                }));
            });
            t.IsBackground = true;
            t.Start();
        }
        // Нажата кнопка просмотра правил входящих сообщений
        private void btInboxMessageRule_Click(object sender, RoutedEventArgs e)
        {
            InboxMessageRule _dwIMR = new InboxMessageRule(loadedData.IndoxMessageRule);
            _dwIMR.ShowDialog();
        }
        // Нажата кнопка просмотра мобильных устройств
        private void btMobileDeviceInfo_Click(object sender, RoutedEventArgs e)
        {
            MobileDevice _dwMD = new MobileDevice(loadedData.MobileDevice);
            _dwMD.ShowDialog();
        }
        // Нажата кнопка просмотра автоответа
        private void btAutoReplyInfo_Click(object sender, RoutedEventArgs e)
        {
            AutoReply _dwAR = new AutoReply(loadedData.AutoReplay);
            _dwAR.ShowDialog();
        }
        // Нажата кнопка просмотра ошибок
        private void btShowErrors_Click(object sender, RoutedEventArgs e)
        {
            ShowErrors _dwSE = new ShowErrors(_errorMessages);
            _dwSE.ShowDialog();
        }
        // В поле выбранного пользователя нажата клавиша
        private void loginUser_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                isEnableForm(false);
                string searchStr = loginUser.Text;
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
                                isEnableForm(true);
                                btSelectUser_Click("loginUser", new RoutedEventArgs());
                            }));
                        }
                        else
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                loginUser.Text = (string)searchResults[0].Properties["sAMAccountName"][0];
                                _userIsSelected = true;
                                loginUser.FontWeight = FontWeights.Bold;
                                loginUser.TextDecorations = TextDecorations.Underline;
                                isEnableForm(true);
                            }));
                        }
                    }
                    else
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            isEnableForm(true);
                            btSelectUser_Click("loginUser", new RoutedEventArgs());
                        }));
                    }
                });
                t.IsBackground = true;
                t.Start();
            }
        }
        // Изменено значение поля выбранного пользователя
        private void loginUser_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_userIsSelected)
            {
                _userIsSelected = false;
                loginUser.FontWeight = FontWeights.Normal;
                loginUser.TextDecorations = null;
            }
        }
    }
}
