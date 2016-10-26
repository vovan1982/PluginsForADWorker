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

namespace AutoReplyAndForwarding
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
        private bool _forwardingMailIsSelected; // Признак того что эл. адрес пересылки выбран

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
            startDate.SelectedDate = DateTime.Today;
            endDate.SelectedDate = DateTime.Today.AddDays(1);
            initSelectedTime(ref startTime);
            initSelectedTime(ref endTime);
            if (startTime.Items.Count > 0)
                startTime.SelectedIndex = 0;
            if (endTime.Items.Count > 0)
                endTime.SelectedIndex = 0;
        }

        // Нажата кнопка выбора сотрудника
        private void btSelectUser_Click(object sender, RoutedEventArgs e)
        {
            SelectUser _dwSU = null;
            if (sender is string)
            {
                if ((string)sender == "loginUserForChange")
                {
                    _dwSU = new SelectUser(_ADSession, _principalContext, loginUserForChange.Text);
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
                loginUserForChange.Text = _dwSU.SelectedUser.Login;
                _userIsSelected = true;
                loginUserForChange.FontWeight = FontWeights.Bold;
                loginUserForChange.TextDecorations = TextDecorations.Underline;
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
        // Нажата кнопка установки автоответа и пересылки сообщений
        private void btSetAutoreplyAndForwarding_Click(object sender, RoutedEventArgs e)
        {
            #region Проверка и инициализация входных данных
            if (!_userIsSelected)
            {
                MessageBox.Show("Не выбран пользователь!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            bool scheduled = false;
            bool forward = false;
            bool autoreply = false;
            bool withOutExternalAutoreply = false;
            string errorMesg = "";
            string loginUser = loginUserForChange.Text;
            DateTime startDateTime = new DateTime();
            DateTime endDateTime = new DateTime();
            string mailAddressForwarding = "";
            string internalMessageData = "";
            string externalMessageData = "";
            if(setScheduled.IsChecked == true)
            {
                scheduled = true;
                startDateTime = new DateTime(
                    startDate.SelectedDate.Value.Year,
                    startDate.SelectedDate.Value.Month,
                    startDate.SelectedDate.Value.Day,
                    TimeSpan.Parse(startTime.SelectedItem.ToString() + ":00").Hours,
                    TimeSpan.Parse(startTime.SelectedItem.ToString() + ":00").Minutes,
                    TimeSpan.Parse(startTime.SelectedItem.ToString() + ":00").Seconds);
                endDateTime = new DateTime(
                    endDate.SelectedDate.Value.Year,
                    endDate.SelectedDate.Value.Month,
                    endDate.SelectedDate.Value.Day,
                    TimeSpan.Parse(endTime.SelectedItem.ToString() + ":00").Hours,
                    TimeSpan.Parse(endTime.SelectedItem.ToString() + ":00").Minutes,
                    TimeSpan.Parse(endTime.SelectedItem.ToString() + ":00").Seconds);
                if (startDateTime >= endDateTime)
                {
                    MessageBox.Show("Дата и время начала периода автоответа не могут быть больше или равны конечной дате и времени!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            if (!string.IsNullOrWhiteSpace(internalMessage.Text))
            {
                internalMessageData = internalMessage.Text.Trim().Replace("\r\n", "<br>");
                externalMessageData = externalMessage.Text.Trim().Replace("\r\n", "<br>");
                autoreply = true;
                if (string.IsNullOrWhiteSpace(externalMessage.Text))
                {
                    withOutExternalAutoreply = true;
                }
            }
            if (setForward.IsChecked == true)
            {
                forward = true;
                if (!_forwardingMailIsSelected)
                {
                    MessageBox.Show("Не указан эл.адрес для пересылки сообщений!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                mailAddressForwarding = mailForwarding.Text;
            }
            if (string.IsNullOrWhiteSpace(_pathToConfig))
            {
                MessageBox.Show("Отсутствует конфигурационный файл плагина!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (!autoreply && !forward)
            {
                MessageBox.Show("Не указаны параметры для установки автоответа или пересылки!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            #endregion

            if (autoreply || forward)
            {
                Thread t = new Thread(() =>
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _processRun.Clear();
                        _processRun.Add(new ProcessRunData { Text = "---- Установка автоответа и переадресации пользователю " + loginUser + " ----" });
                        isEnableForm(false);
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
                            itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Ошибка чтения конфигурационного файла:\r\n" + errorMesg;
                            isEnableForm(true);
                        }));
                        return;
                    }
                    if (!configData.ContainsKey("mailserver"))
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Отсутсвуют требуемые параметры в конфигурационном файле!!!";
                            isEnableForm(true);
                        }));
                        return;
                    }
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
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

                    #region Внесение изменений

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
                            itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Ошибка подключения к Exchange Server:\r\n" + exp.Message;
                            isEnableForm(true);
                        }));
                        return;
                    }
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Подключение к Exchange Server";
                    }));
                    #endregion

                    if (autoreply)
                    {
                        #region Установка автоответа
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            _processRun.Add(new ProcessRunData { Text = "Установка автоответа..." });
                            itemProcess = _processRun[_processRun.Count - 1];
                        }));
                        try
                        {
                            command.AddCommand("Set-MailboxAutoReplyConfiguration");
                            command.AddParameter("Identity", loginUser);
                            if (scheduled)
                            {
                                command.AddParameter("AutoReplyState", "Scheduled");
                                command.AddParameter("StartTime", startDateTime.ToString("M/dd/yyyy HH:mm:ss").Replace('.', '/'));
                                command.AddParameter("EndTime", endDateTime.ToString("M/dd/yyyy HH:mm:ss").Replace('.', '/'));
                            }
                            else
                            {
                                command.AddParameter("AutoReplyState", "Enabled");
                            }
                            if (withOutExternalAutoreply)
                                command.AddParameter("ExternalAudience", "None");
                            else
                                command.AddParameter("ExternalAudience", "All");
                            command.AddParameter("ExternalMessage", externalMessageData);
                            command.AddParameter("InternalMessage", internalMessageData);
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
                                    string errorMsg = "Ошибка установки автоответа:\r\n" + powershell.Streams.Error[i].ToString();
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                        itemProcess.Text += errorMsg;
                                    }));
                                }
                                runspace.Dispose();
                                runspace = null;
                                powershell.Dispose();
                                powershell = null;
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    isEnableForm(true);
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
                                itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = "Ошибка установки автоответа:\r\n" + exp.Message;
                                isEnableForm(true);
                            }));
                            return;
                        }
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Установка автоответа";
                        }));
                        #endregion
                    }

                    if (forward)
                    {
                        #region Проверка адреса пересылки сообщений
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            _processRun.Add(new ProcessRunData { Text = "Проверка адреса пересылки сообщений..." });
                            itemProcess = _processRun[_processRun.Count - 1];
                        }));
                        try
                        {
                            command.AddCommand("Get-Mailbox");
                            command.AddParameter("Identity", mailAddressForwarding);
                            powershell.Commands = command;
                            var result = powershell.Invoke();
                            command.Clear();
                            if (result.Count <= 0)
                            {
                                runspace.Dispose();
                                runspace = null;
                                powershell.Dispose();
                                powershell = null;
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                    itemProcess.Text += "Ошибка проверки адреса пересылки сообщений:\r\nАдрес пересылки не найден на сервере Exchange!!!";
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
                                itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = "Ошибка проверки адреса пересылки сообщений:\r\n" + exp.Message;
                                isEnableForm(true);
                            }));
                            return;
                        }
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Проверка адреса пересылки сообщений";
                        }));
                        #endregion

                        #region Установка пересылки входящих сообщений
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            _processRun.Add(new ProcessRunData { Text = "Установка пересылки входящих сообщений..." });
                            itemProcess = _processRun[_processRun.Count - 1];
                        }));
                        try
                        {
                            command.AddCommand("Set-Mailbox");
                            command.AddParameter("Identity", loginUser);
                            command.AddParameter("DeliverToMailboxAndForward", true);
                            command.AddParameter("ForwardingAddress", mailAddressForwarding);
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
                                    string errorMsg = "Ошибка установки пересылки входящих сообщений:\r\n" + powershell.Streams.Error[i].ToString();
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                        itemProcess.Text += errorMsg;
                                    }));
                                }
                                runspace.Dispose();
                                runspace = null;
                                powershell.Dispose();
                                powershell = null;
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    isEnableForm(true);
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
                                itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = "Ошибка установки пересылки входящих сообщений:\r\n" + exp.Message;
                                isEnableForm(true);
                            }));
                            return;
                        }
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/AutoReplyAndForwarding;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Установка пересылки входящих сообщений";
                        }));
                        #endregion
                    }

                    #endregion

                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    command.Clear();
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _processRun.Add(new ProcessRunData { Text = "---- Процесс завершён! ----" });
                        loginUserForChange.Text = "";
                        _userIsSelected = false;
                        loginUserForChange.FontWeight = FontWeights.Normal;
                        loginUserForChange.TextDecorations = null;
                        startDate.SelectedDate = DateTime.Today;
                        endDate.SelectedDate = DateTime.Today.AddDays(1);
                        startTime.SelectedIndex = 0;
                        endTime.SelectedIndex = 0;
                        internalMessage.Text = "";
                        externalMessage.Text = "";
                        mailForwarding.Text = "";
                        _forwardingMailIsSelected = false;
                        mailForwarding.FontWeight = FontWeights.Normal;
                        mailForwarding.TextDecorations = null;
                        setForward.IsChecked = false;
                        isEnableForm(true);
                    }));
                });
                t.IsBackground = true;
                t.Start();
            }
            else
            {
                if (scheduled)
                {
                    MessageBox.Show("Не указаны сообщения для отправки в пределах и за пределы организации!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                    MessageBox.Show("Не указаны значения для установки!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Нажат выбор интервала автоответа
        private void setScheduled_Click(object sender, RoutedEventArgs e)
        {
            if (setScheduled.IsChecked == true)
            {
                labelStartDateTime.IsEnabled = true;
                startDate.IsEnabled = true;
                startTime.IsEnabled = true;
                labelEndDateTime.IsEnabled = true;
                endDate.IsEnabled = true;
                endTime.IsEnabled = true;
            }
            else
            {
                labelStartDateTime.IsEnabled = false;
                startDate.IsEnabled = false;
                startDate.SelectedDate = DateTime.Today;
                startTime.IsEnabled = false;
                if (startTime.Items.Count > 0)
                    startTime.SelectedIndex = 0;
                labelEndDateTime.IsEnabled = false;
                endDate.IsEnabled = false;
                endDate.SelectedDate = DateTime.Today.AddDays(1);
                endTime.IsEnabled = false;
                if (endTime.Items.Count > 0)
                    endTime.SelectedIndex = 0;
            }
        }
        // Нажат выбор пересылки сообщений
        private void setForward_Click(object sender, RoutedEventArgs e)
        {
            if (setForward.IsChecked == true)
            {
                mailForwarding.IsEnabled = true;
                btSelectUserForward.IsEnabled = true;
            }
            else
            {
                mailForwarding.IsEnabled = false;
                mailForwarding.Text = "";
                btSelectUserForward.IsEnabled = false;
            }
        }
        // Инициализация комбобокса времени
        private void initSelectedTime(ref ComboBox init)
        {
            DateTime times = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0);
            for (int i = 0; i < 48; i++)
            {
                string timestr = times.ToString("H:mm");
                init.Items.Add(timestr);
                times = times.AddMinutes(30.0);
            }
        }
        // Нажата кнопка выбора адреса пересылки
        private void btSelectUserForward_Click(object sender, RoutedEventArgs e)
        {
            SelectUser _dwSU = null;
            if (sender is string)
            {
                if ((string)sender == "mailForwarding")
                {
                    _dwSU = new SelectUser(_ADSession, _principalContext, mailForwarding.Text);
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
                mailForwarding.Text = _dwSU.SelectedUser.Mail;
                _forwardingMailIsSelected = true;
                mailForwarding.FontWeight = FontWeights.Bold;
                mailForwarding.TextDecorations = TextDecorations.Underline;
            }
        }
        // Деактивация формы
        private void isEnableForm(bool enable)
        {
            loginUserForChange.IsEnabled = enable;
            btSelectUser.IsEnabled = enable;
            setScheduled.IsEnabled = enable;
            if (setScheduled.IsChecked == true && enable)
            {
                labelStartDateTime.IsEnabled = enable;
                labelStartDateTime.IsEnabled = enable;
                startDate.IsEnabled = enable;
                startTime.IsEnabled = enable;
                labelEndDateTime.IsEnabled = enable;
                endDate.IsEnabled = enable;
                endTime.IsEnabled = enable;
            }
            else if (setScheduled.IsChecked == false && enable)
            {
                labelStartDateTime.IsEnabled = !enable;
                labelStartDateTime.IsEnabled = !enable;
                startDate.IsEnabled = !enable;
                startTime.IsEnabled = !enable;
                labelEndDateTime.IsEnabled = !enable;
                endDate.IsEnabled = !enable;
                endTime.IsEnabled = !enable;
            }
            else
            {
                labelStartDateTime.IsEnabled = enable;
                labelStartDateTime.IsEnabled = enable;
                startDate.IsEnabled = enable;
                startTime.IsEnabled = enable;
                labelEndDateTime.IsEnabled = enable;
                endDate.IsEnabled = enable;
                endTime.IsEnabled = enable;
            }

            internalMessage.IsEnabled = enable;
            externalMessage.IsEnabled = enable;
            setForward.IsEnabled = enable;
            if (setForward.IsChecked == true && enable)
            {
                mailForwarding.IsEnabled = enable;
                btSelectUserForward.IsEnabled = enable;
            }
            else if (setForward.IsChecked == false && enable)
            {
                mailForwarding.IsEnabled = !enable;
                btSelectUserForward.IsEnabled = !enable;
            }
            else
            {
                mailForwarding.IsEnabled = enable;
                btSelectUserForward.IsEnabled = enable;
            }

            btSetAutoreplyAndForwarding.IsEnabled = enable;
        }
        // В поле выбранного пользователя нажата клавиша
        private void loginUserForChange_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                isEnableForm(false);
                string searchStr = loginUserForChange.Text;
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
                                btSelectUser_Click("loginUserForChange", new RoutedEventArgs());
                            }));
                        }
                        else
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                loginUserForChange.Text = (string)searchResults[0].Properties["sAMAccountName"][0];
                                _userIsSelected = true;
                                loginUserForChange.FontWeight = FontWeights.Bold;
                                loginUserForChange.TextDecorations = TextDecorations.Underline;
                                isEnableForm(true);
                            }));
                        }
                    }
                    else
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            isEnableForm(true);
                            btSelectUser_Click("loginUserForChange", new RoutedEventArgs());
                        }));
                    }
                });
                t.IsBackground = true;
                t.Start();
            }
        }
        // Изменено значение поля выбранного пользователя
        private void loginUserForChange_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_userIsSelected)
            {
                _userIsSelected = false;
                loginUserForChange.FontWeight = FontWeights.Normal;
                loginUserForChange.TextDecorations = null;
            }
        }
        // В поле адреса пересылки нажата клавиша
        private void mailForwarding_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                isEnableForm(false);
                string searchStr = mailForwarding.Text;
                Thread t = new Thread(() =>
                {
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(!(objectClass=computer))(|(cn=*{0}*)(sn=*{0}*)(givenName=*{0}*)(sAMAccountName=*{0}*)(mail={0})))", searchStr.Trim());
                    dirSearcher.PropertiesToLoad.Add("mail");
                    SearchResultCollection searchResults = dirSearcher.FindAll();
                    if (searchResults.Count > 0)
                    {
                        if (searchResults.Count > 1)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                isEnableForm(true);
                                btSelectUserForward_Click("mailForwarding", new RoutedEventArgs());
                            }));
                        }
                        else
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                mailForwarding.Text = (string)searchResults[0].Properties["mail"][0];
                                _forwardingMailIsSelected = true;
                                mailForwarding.FontWeight = FontWeights.Bold;
                                mailForwarding.TextDecorations = TextDecorations.Underline;
                                isEnableForm(true);
                            }));
                        }
                    }
                    else
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            isEnableForm(true);
                            btSelectUserForward_Click("mailForwarding", new RoutedEventArgs());
                        }));
                    }
                });
                t.IsBackground = true;
                t.Start();
            }
        }
        // Изменено значение поля адреса пересылки
        private void mailForwarding_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_forwardingMailIsSelected)
            {
                _forwardingMailIsSelected = false;
                mailForwarding.FontWeight = FontWeights.Normal;
                mailForwarding.TextDecorations = null;
            }
        }
    }
}
