using CreateUsersPPK.Model;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Data;
using CreateUsersPPK.Converters;
using System.Text.RegularExpressions;

namespace CreateUsersPPK
{
    /// <summary>
    /// Логика взаимодействия для UserControl1.xaml
    /// </summary>
    public partial class ControlPlMain : UserControl
    {
        #region Поля
        private string _login; // Логин для соединения с почтовым сервером
        private string _pathToConfig; // Полный путь к конфигурационному файлу плагина
        private string _pathToPluginDir; // Полный путь к папке плагина
        private SecureString _password; // Пароль для соединения с почтовым сервером
        private DirectoryEntry _ADSession; // Сессия с АД
        private PrincipalContext _principalContext; // Контекст с АД
        private Dictionary<string, string> configData; // Данные конфигурационного файла
        private ObservableCollection<ProcessRunData> _processRun; // Лог
        private ObservableCollection<Cities> _cities; // Список городов доступных для выбора
        private ObservableCollection<string> _posts; // Список должностей доступных для выбора
        private ObservableCollection<string> _departments; // Список отделов доступных для выбора
        private ObservableCollection<string> _mailDataBase; // Список почтовых баз данных доступных для выбора
        private ObservableCollection<string> _activeSyncPolicy; // Список политик ActiveSync доступных для выбора
        private ObservableCollection<string> _owaPolicy; // Список политик OWA доступных для выбора
        private ObservableCollection<string> _templates; // Список шаблонов доступных для выбора
        private ObservableCollection<string> _groups; // Список групп в которых должен состоять пользователь
        private bool isCheckedLogin; // Логин проверен на соответствие
        private bool isCheckedNameInAD; // Имя в АД проверено на соответствие
        private bool isCheckedMail; // Почта проверена на соответствие
        #endregion

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
            _processRun = new ObservableCollection<ProcessRunData>();
            _groups = new ObservableCollection<string>();
            goupListToAdd.ItemsSource = _groups;
            lbVersion.Content = lbVersion.Content + " " + plug.Version;
            _pathToPluginDir = _pathToConfig.Substring(0,_pathToConfig.LastIndexOf("\\"));
            isCheckedLogin = false;
            isCheckedNameInAD = false;
            isCheckedMail = false;
            LoadConfigData();
        }

        #region События
        // Нажата кнопка выбора OU пользователя
        private void btSelectOU_Click(object sender, RoutedEventArgs e)
        {
            DialogWindows.SelectOU _dwMUIAD = new DialogWindows.SelectOU(_ADSession);
            _dwMUIAD.Owner = Application.Current.MainWindow;
            bool? result = _dwMUIAD.ShowDialog();
            if (result == true)
                placeInAD.Text = _dwMUIAD.SelectedOU;
        }
        // Поле Фамилия потеряло фокус
        private void surname_LostFocus(object sender, RoutedEventArgs e)
        {
            login.Text = CreateLogin();
            nameInAD.Text = CreateNameInAD();
            if (nameInAD.Text.Length >= 64)
                nameInAD.IsReadOnly = false;
            mail.Text = CreateMail();
        }
        // Поле Имя потеряло фокус
        private void name_LostFocus(object sender, RoutedEventArgs e)
        {
            login.Text = CreateLogin();
            nameInAD.Text = CreateNameInAD();
            if (nameInAD.Text.Length >= 64)
                nameInAD.IsReadOnly = false;
            mail.Text = CreateMail();
        }
        // Поле отчество потеряло фокус
        private void middlename_LostFocus(object sender, RoutedEventArgs e)
        {
            login.Text = CreateLogin();
            nameInAD.Text = CreateNameInAD();
            if (nameInAD.Text.Length >= 64)
                nameInAD.IsReadOnly = false;
            mail.Text = CreateMail();
        }
        // Нажата кнопка чтения конфиг файла
        private void btReloadConfig_Click(object sender, RoutedEventArgs e)
        {
            LoadConfigData();
        }
        // Нажата кнопка просмотра отчёта
        private void btReport_Click(object sender, RoutedEventArgs e)
        {
            DialogWindows.Report _dwR = new DialogWindows.Report(_processRun);
            _dwR.Owner = Application.Current.MainWindow;
            _dwR.ShowDialog();
        }
        // Изменен город
        private void city_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (city.SelectedItem != null)
                    adress.Text = ((Cities)city.SelectedItem).Adress;
            }
            catch (Exception exp)
            {
                MessageBox.Show("Не удалось установить адрес:\r\n" + exp.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Изменен шаблон
        private void templates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (templates.SelectedItem != null)
            {
                btLoadTemplate.IsEnabled = true;
            }
            else
            {
                btLoadTemplate.IsEnabled = false;
            }
        }
        // Нажат checkBox ActiveSync
        private void chActiveSync_Click(object sender, RoutedEventArgs e)
        {
            if (chActiveSync.IsChecked == true)
            {
                labelActiveSyncPolicy.IsEnabled = true;
                activeSyncPolicy.IsEnabled = true;
            }
            else
            {
                labelActiveSyncPolicy.IsEnabled = false;
                activeSyncPolicy.IsEnabled = false;
            }
        }
        // Нажат checkBox OWA
        private void chOWA_Click(object sender, RoutedEventArgs e)
        {
            if (chOWA.IsChecked == true)
            {
                labelOwaPolicy.IsEnabled = true;
                owaPolicy.IsEnabled = true;
            }
            else
            {
                labelOwaPolicy.IsEnabled = false;
                owaPolicy.IsEnabled = false;
            }
        }
        // Выполнен перереход с поля Город
        private void city_LostFocus(object sender, RoutedEventArgs e)
        {
            if (city.SelectedItem == null && !string.IsNullOrWhiteSpace(city.Text))
            {
                    adress.Text = "";
                    city.SelectedIndex = -1;
                    city.Text = "";
                    MessageBox.Show("Указанного города нет в списке выбора!!!\r\nДобавьте город через форму добавления или выберите существующий.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Нажата кнопка добавления групп
        private void btAddGroups_Click(object sender, RoutedEventArgs e)
        {
            DialogWindows.SelectGroups _dwSG = new DialogWindows.SelectGroups(_ADSession);
            _dwSG.Owner = Application.Current.MainWindow;
            bool? result = _dwSG.ShowDialog();
            if (result == true)
            {
                if (_groups.Count > 0)
                {
                    bool isExist = false;
                    for (int i = 0; i < _dwSG.SelectedGroups.Count; i++)
                    {
                        for (int j = 0; j < _groups.Count; j++)
                        {
                            if (_groups[j] == _dwSG.SelectedGroups[i].Name)
                            {
                                isExist = true;
                                break;
                            }
                        }
                        if(!isExist)
                            _groups.Add(_dwSG.SelectedGroups[i].Name);
                    }
                }
                else
                {
                    for (int i = 0; i < _dwSG.SelectedGroups.Count; i++)
                    {
                        _groups.Add(_dwSG.SelectedGroups[i].Name);
                    }
                }
                foreach (var column in goupListToAddGridView.Columns)
                {
                    column.Width = column.ActualWidth;
                    column.Width = Double.NaN;
                }
            }
        }
        // Нажата кнопка удаления групп
        private void btDeleteGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                while (goupListToAdd.SelectedItems.Count > 0)
                {
                    _groups.Remove((string)goupListToAdd.SelectedItems[0]);
                }
                if (goupListToAdd.Items.Count > 0)
                {
                    goupListToAdd.SelectedIndex = 0;
                    goupListToAdd.Focus();
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Нажата кнопка добавления города
        private void btAddCity_Click(object sender, RoutedEventArgs e)
        {
            DialogWindows.DialogAddEditCity _dwDAEC = new DialogWindows.DialogAddEditCity(_pathToPluginDir, _cities.ToList(), _ADSession, FormModes.Mode.add);
            _dwDAEC.Owner = Application.Current.MainWindow;
            bool? result = _dwDAEC.ShowDialog();
            if (result == true)
            {
                string[] citiesArr = _dwDAEC.NewCityToAdd.Split('\n');
                for (int i = 0; i < citiesArr.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(citiesArr[i]))
                    {
                        List<string> cityGroups = new List<string>();
                        string[] city = citiesArr[i].TrimEnd('\r').Split(';');
                        string[] cityGroup = city[3].Split(',');
                        if (cityGroup.Length > 0)
                            for (int j = 0; j < cityGroup.Length; j++)
                            {
                                cityGroups.Add(cityGroup[j]);
                            }
                        _cities.Add(new Cities { DisplayName = city[0], Name = city[1], Adress = city[2], Groups = cityGroups });
                    }
                }
            }

        }
        // Нажата кнопка редактирования города
        private void btEditCity_Click(object sender, RoutedEventArgs e)
        {
            if (city.SelectedIndex >= 0)
            {
                DialogWindows.DialogAddEditCity _dwDAEC = new DialogWindows.DialogAddEditCity(_pathToPluginDir, _cities.ToList(), _ADSession, FormModes.Mode.edit, city.SelectedIndex);
                _dwDAEC.Owner = Application.Current.MainWindow;
                bool? result = _dwDAEC.ShowDialog();
                if (result == true)
                {
                    int index = city.SelectedIndex;
                    if (!string.IsNullOrWhiteSpace(_dwDAEC.NewCityToAdd))
                    {
                        List<string> cityGroups = new List<string>();
                        string[] newCity = _dwDAEC.NewCityToAdd.Split(';');
                        string[] cityGroup = newCity[3].Split(',');
                        if (cityGroup.Length > 0)
                            for (int j = 0; j < cityGroup.Length; j++)
                            {
                                if (string.IsNullOrWhiteSpace(cityGroup[j]))
                                    cityGroups.Add(cityGroup[j]);
                            }
                        _cities[index].DisplayName = newCity[0];
                        _cities[index].Name = newCity[1];
                        _cities[index].Adress = newCity[2];
                        _cities[index].Groups = cityGroups;
                        adress.Text = newCity[2];
                        city.Text = newCity[0];
                        city.SelectedIndex = index;
                    }
                }
            }
        }
        // Нажата кнопка обновления городов
        private void btUpdateCity_Click(object sender, RoutedEventArgs e)
        {
            statusText.Content = "Загрузка городов из конфигурационного файла...";
            IsEnableForm(false);
            _processRun.Clear();
            string dataFull = "";
            _cities = new ObservableCollection<Cities>();
            Thread t = new Thread(() =>
            {
                ProcessRunData itemProcess = new ProcessRunData();

                #region Загрузка списка городов для выбора
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Загрузка городов из конфигурационного файла..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (File.Exists(_pathToPluginDir + "\\cities.dat"))
                    {
                        using (StreamReader sr = new StreamReader(_pathToPluginDir + "\\cities.dat", Encoding.GetEncoding("windows-1251")))
                        {
                            dataFull = sr.ReadToEnd();
                            string[] configArr = dataFull.Split('\n');
                            for (int i = 0; i < configArr.Length; i++)
                            {
                                if (!string.IsNullOrWhiteSpace(configArr[i]))
                                {
                                    List<string> cityGroups = new List<string>();
                                    string[] city = configArr[i].TrimEnd('\r').Split(';');
                                    string[] cityGroup = city[3].Split(',');
                                    if (cityGroup.Length > 0)
                                        for (int j = 0; j < cityGroup.Length; j++)
                                        {
                                            if (!string.IsNullOrWhiteSpace(cityGroup[j]))
                                                cityGroups.Add(cityGroup[j]);
                                        }
                                    _cities.Add(new Cities { DisplayName = city[0], Name = city[1], Adress = city[2], Groups = cityGroups });
                                }
                            }
                        }
                    }
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка загрузки городов из конфигурационного файла:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    city.ItemsSource = _cities;
                    city.DisplayMemberPath = "DisplayName";
                    adress.Text = "";
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Загрузка городов из конфигурационного файла завершена";
                    _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    statusText.Content = "";
                    IsEnableForm(true);
                }));
                #endregion
            });
            t.IsBackground = true;
            t.Start();
        }
        // Нажата кнопка сохранения должности
        private void btSavePost_Click(object sender, RoutedEventArgs e)
        {
            if (post.SelectedItem == null)
            {
                if (!string.IsNullOrWhiteSpace(post.Text))
                {
                    try
                    {
                        #region Запись новых данных в файл
                        if (File.Exists(_pathToPluginDir + "\\posts.dat"))
                        {
                            StreamReader sr = new StreamReader(_pathToPluginDir + "\\posts.dat", Encoding.GetEncoding("windows-1251"));
                            string fullData = sr.ReadToEnd();
                            sr.Close();
                            if (fullData.EndsWith("\r\n"))
                            {
                                File.AppendAllText(_pathToPluginDir + "\\posts.dat", post.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                            }
                            else if (string.IsNullOrWhiteSpace(fullData))
                            {
                                File.AppendAllText(_pathToPluginDir + "\\posts.dat", post.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                            }
                            else
                            {
                                File.AppendAllText(_pathToPluginDir + "\\posts.dat", "\r\n" + post.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                            }
                        }
                        else
                        {
                            File.WriteAllText(_pathToPluginDir + "\\posts.dat", post.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                        }
                        #endregion
                        _posts.Add(post.Text);
                        post.SelectedIndex = _posts.Count - 1;
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("Не удалось сохранить должность:\r\n" + exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
        // Нажата кнопка сохранения отдела
        private void btSaveDepartment_Click(object sender, RoutedEventArgs e)
        {
            if (department.SelectedItem == null)
            {
                if (!string.IsNullOrWhiteSpace(department.Text))
                {
                    try
                    {
                        #region Запись новых данных в файл
                        if (File.Exists(_pathToPluginDir + "\\departments.dat"))
                        {
                            StreamReader sr = new StreamReader(_pathToPluginDir + "\\departments.dat", Encoding.GetEncoding("windows-1251"));
                            string fullData = sr.ReadToEnd();
                            sr.Close();
                            if (fullData.EndsWith("\r\n"))
                            {
                                File.AppendAllText(_pathToPluginDir + "\\departments.dat", department.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                            }
                            else if (string.IsNullOrWhiteSpace(fullData))
                            {
                                File.AppendAllText(_pathToPluginDir + "\\departments.dat", department.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                            }
                            else
                            {
                                File.AppendAllText(_pathToPluginDir + "\\departments.dat", "\r\n" + department.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                            }
                        }
                        else
                        {
                            File.WriteAllText(_pathToPluginDir + "\\departments.dat", department.Text + "\r\n", Encoding.GetEncoding("windows-1251"));
                        }
                        #endregion
                        _departments.Add(department.Text);
                        department.SelectedIndex = _departments.Count - 1;
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("Не удалось сохранить отдел:\r\n" + exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
        // Нажата кнопка загрузки шаблона
        private void btLoadTemplate_Click(object sender, RoutedEventArgs e)
        {
            _groups.Clear();
            try
            {
                string dataFull = "";
                using (StreamReader sr = new StreamReader(_pathToPluginDir + "\\templates\\" + templates.SelectedItem.ToString() + ".txt", Encoding.GetEncoding("windows-1251")))
                {
                    dataFull = sr.ReadToEnd();
                    string[] groupArr = dataFull.Split('\n');
                    for (int i = 0; i < groupArr.Length; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(groupArr[i]))
                        {
                            _groups.Add(groupArr[i].TrimEnd('\r'));
                        }
                    }
                }
                foreach (var column in goupListToAddGridView.Columns)
                {
                    column.Width = column.ActualWidth;
                    column.Width = Double.NaN;
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show("Не удалось загрузить шаблон:\r\n" + exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Нажата кнопка создания шаблона
        private void btAddTemplate_Click(object sender, RoutedEventArgs e)
        {
            DialogWindows.DialogAddEditTemplate _dwDAET = new DialogWindows.DialogAddEditTemplate(_pathToPluginDir, _templates.ToList(), _ADSession, FormModes.Mode.add);
            _dwDAET.Owner = Application.Current.MainWindow;
            bool? result = _dwDAET.ShowDialog();
            if (result == true)
            {
                string[] templateArr = _dwDAET.NewTemplateToAdd.Split('\n');
                for (int i = 0; i < templateArr.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(templateArr[i]))
                    {
                        _templates.Add(templateArr[i].TrimEnd('\r'));
                    }
                }
            }
        }
        // Нажата кнопка редактирования шаблона
        private void btEditTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (templates.SelectedIndex >= 0)
            {
                DialogWindows.DialogAddEditTemplate _dwDAET = new DialogWindows.DialogAddEditTemplate(_pathToPluginDir, _templates.ToList(), _ADSession, FormModes.Mode.edit, templates.SelectedIndex);
                _dwDAET.Owner = Application.Current.MainWindow;
                bool? result = _dwDAET.ShowDialog();
                if (result == true)
                {
                    int index = templates.SelectedIndex;
                    if (!string.IsNullOrWhiteSpace(_dwDAET.NewTemplateToAdd))
                    {
                        _templates[templates.SelectedIndex] = _dwDAET.NewTemplateToAdd;
                        templates.SelectedIndex = index;
                    }
                }
            }
        }
        // Нажата кнопка обновления списка шаблонов
        private void btUpdateTemplate_Click(object sender, RoutedEventArgs e)
        {
            _templates.Clear();
            try
            {
                if (Directory.Exists(_pathToPluginDir + "\\templates"))
                {
                    string[] arrTemplates = Directory.GetFiles(_pathToPluginDir + "\\templates");
                    if (arrTemplates.Length > 0)
                    {
                        for (int j = 0; j < arrTemplates.Length; j++)
                        {
                            _templates.Add(arrTemplates[j].Substring(arrTemplates[j].LastIndexOf("\\")).TrimStart('\\').TrimEnd(new char[] { '.', 't', 'x', 't' }));
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show("Не удалось сформировать список шаблонов:\r\n" + exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Изменен логин
        private void login_TextChanged(object sender, TextChangedEventArgs e)
        {
            isCheckedLogin = false;
            login.Foreground = new SolidColorBrush(Colors.Black);
        }
        // Изменено Имя в АД
        private void nameInAD_TextChanged(object sender, TextChangedEventArgs e)
        {
            isCheckedNameInAD = false;
            nameInAD.Foreground = new SolidColorBrush(Colors.Black);
        }
        // Изменен эл. адрес
        private void mail_TextChanged(object sender, TextChangedEventArgs e)
        {
            isCheckedMail = false;
            mail.Foreground = new SolidColorBrush(Colors.Black);
        }
        // Именена OU
        private void placeInAD_TextChanged(object sender, TextChangedEventArgs e)
        {
            isCheckedNameInAD = false;
            nameInAD.Foreground = new SolidColorBrush(Colors.Black);
        }
        // Нажата кнопка проверки логина
        private void checkLogin_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(login.Text))
            {
                MessageBox.Show("Логин не указан!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            statusText.Content = "Проверка логина на соответствие требованиям...";
            string loginData = login.Text;
            IsEnableForm(false);
            _processRun.Clear();
            lstpop.Items.Clear();
            _processRun.Add(new ProcessRunData { Text = "----Проверка логина на соответствие требованиям----" });
            Thread t = new Thread(() =>
            {
                ProcessRunData itemProcess = new ProcessRunData();
                #region Проверка длинны логина
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка длинны логина..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                if (loginData.Length > 20)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Длинна логина превышает допустимую длинну в 20 символов";
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        lstpop.Items.Add("Длинна логина превышает допустимую длинну в 20 символов!!!");
                        login.Foreground = new SolidColorBrush(Colors.Red);
                        popLogin.IsOpen = true;
                        statusText.Content = "";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка длинны логина выполнена успешно";
                }));
                #endregion
                #region Проверка наличия логина в АД
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка наличия логина в АД..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("(&(objectClass=user)(!(objectClass=computer))(sAMAccountName={0}))", loginData);
                    SearchResult searchResults = dirSearcher.FindOne();
                    if (searchResults != null)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Такой логин уже присутствует в АД!!!";
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            lstpop.Items.Add("Такой логин уже присутствует в АД!!!");
                            lstpop.Items.Add("Измените логин.");
                            login.Foreground = new SolidColorBrush(Colors.Red);
                            popLogin.IsOpen = true;
                            statusText.Content = "";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка проверки наличия логина в АД:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе проверки наличия логина в АД возникли проблемы, детальную информацию смотрите в отчете.";
                        login.Foreground = new SolidColorBrush(Colors.Red);
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка наличия логина в АД выполнена успешно";
                }));
                #endregion
                
                isCheckedLogin = true;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    login.Foreground = new SolidColorBrush(Colors.Green);
                    statusText.Content = "";
                    IsEnableForm(true);
                }));
            });
            t.IsBackground = true;
            t.Start();
        }
        // Нажата кнопка проверки имени в АД
        private void checkNameInAD_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(nameInAD.Text))
            {
                MessageBox.Show("Не указано имя в АД!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(placeInAD.Text))
            {
                MessageBox.Show("Не выбрано OU пользователя!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            statusText.Content = "Проверка имени в АД на наличие в указанной OU...";
            string nameInADData = nameInAD.Text;
            string ou = placeInAD.Text;
            IsEnableForm(false);
            _processRun.Clear();
            lstpopname.Items.Clear();
            _processRun.Add(new ProcessRunData { Text = "----Проверка имени в АД на наличие в указанной OU----" });
            Thread t = new Thread(() =>
            {
                ProcessRunData itemProcess = new ProcessRunData();
                #region Проверка имени в АД
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка имени в АД..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("distinguishedName={0}", "CN=" + nameInADData + "," + ou);
                    SearchResult searchResults = dirSearcher.FindOne();
                    if (searchResults != null)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Такое имя уже присутствует в выбранной OU!!!";
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            lstpopname.Items.Add("Такое имя уже присутствует в выбранной OU!!!");
                            lstpopname.Items.Add("Измените имя или выберите другую OU.");
                            nameInAD.Foreground = new SolidColorBrush(Colors.Red);
                            nameInAD.IsReadOnly = false;
                            popNameInAD.IsOpen = true;
                            statusText.Content = "";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка проверки имени в АД на наличие в указанной OU:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе проверки имени в АД на наличие в указанной OU, возникли проблемы, детальную информацию смотрите в отчете.";
                        nameInAD.Foreground = new SolidColorBrush(Colors.Red);
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка имени в АД на наличие в указанной OU выполнена успешно";
                }));
                #endregion
                isCheckedNameInAD = true;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    nameInAD.Foreground = new SolidColorBrush(Colors.Green);
                    statusText.Content = "";
                    IsEnableForm(true);
                }));
            });
            t.IsBackground = true;
            t.Start();
        }
        // Нажата кнопка проверки почты
        private void btCheckMail_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(mail.Text))
            {
                MessageBox.Show("Не указан эл. адрес!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            statusText.Content = "Проверка эл. адреса...";
            string mailData = mail.Text;
            IsEnableForm(false);
            _processRun.Clear();
            lstpopmail.Items.Clear();
            _processRun.Add(new ProcessRunData { Text = "----Проверка эл. адреса----" });
            Thread t = new Thread(() =>
            {
                ProcessRunData itemProcess = new ProcessRunData();

                #region Проверка корректности эл. адреса
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка корректности эл. адреса..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (!Regex.IsMatch(mailData, @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase))
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Не корректный эл. адрес!!!";
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            lstpopmail.Items.Add("Не корректный эл. адрес!!!");
                            lstpopmail.Items.Add("Измените эл. адрес.");
                            mail.Foreground = new SolidColorBrush(Colors.Red);
                            popMail.IsOpen = true;
                            statusText.Content = "";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка проверки корректности эл. адреса:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе проверки эл. адреса возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка корректности эл. адреса выполнена успешно";
                }));
                #endregion

                #region Проверка корректности домена эл. адреса
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка корректности домена эл. адреса..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (mailData.Split('@')[1] != configData["mailprimarydomain"])
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Не корректный домен эл. адреса!!!";
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            lstpopmail.Items.Add("Не корректный домен эл. адреса!!!");
                            lstpopmail.Items.Add("Домен эл. адреса отличается от основного домена указанного");
                            lstpopmail.Items.Add("в конфигурационном файле.");
                            lstpopmail.Items.Add("Укажите в качестве домена эл. адреса: " + configData["mailprimarydomain"]);
                            mail.Foreground = new SolidColorBrush(Colors.Red);
                            popMail.IsOpen = true;
                            statusText.Content = "";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка проверки корректности домена эл. адреса:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе проверки эл. адреса возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка корректности домена эл. адреса выполнена успешно";
                }));
                #endregion
                
                // Настройка параметров подключения к Exchange server
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
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка подключения к Exchange Server:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе проверки эл. адреса возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Подключение к Exchange Server выполнено успешно";
                }));
                #endregion

                #region Проверка эл. адреса в Exchange
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Проверка эл. адреса в Exchange..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    command.AddCommand("Get-Mailbox");
                    command.AddParameter("Identity", mailData);
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (result.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Такой эл. адрес уже существует!!!";
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            lstpopmail.Items.Add("Такой эл. адрес уже существует!!!");
                            lstpopmail.Items.Add("Измените эл. адрес.");
                            mail.Foreground = new SolidColorBrush(Colors.Red);
                            popMail.IsOpen = true;
                            statusText.Content = "";
                            IsEnableForm(true);
                        }));
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        command.Clear();
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка проверки эл. адреса в Exchange:\r\n" + ex.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    }));
                    command.Clear();
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе проверки эл. адреса возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка эл. адреса в Exchange выполнена успешно";
                }));
                #endregion

                // Закрытие соединений с Exchange server
                //============================================================
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
                command.Clear();
                //============================================================
                isCheckedMail = true;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    mail.Foreground = new SolidColorBrush(Colors.Green);
                    statusText.Content = "";
                    IsEnableForm(true);
                    btCheckMail.Focus();
                }));
            });
            t.IsBackground = true;
            t.Start();
        }
        // Нажата кнопка создания учетной записи
        private void btCreateUser_Click(object sender, RoutedEventArgs e)
        {
            NewUser newUserData = new NewUser();
            int CreateDataStage = 0;
            try
            {
                #region Проверка входных значений
                if (!isCheckedLogin || !isCheckedMail || !isCheckedNameInAD)
                {
                    MessageBox.Show("Не все проверки выполнены!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                if (string.IsNullOrWhiteSpace(surname.Text) ||
                    string.IsNullOrWhiteSpace(name.Text) ||
                    string.IsNullOrWhiteSpace(middlename.Text) ||
                    string.IsNullOrWhiteSpace(placeInAD.Text) ||
                    string.IsNullOrWhiteSpace(login.Text) ||
                    string.IsNullOrWhiteSpace(nameInAD.Text) ||
                    string.IsNullOrWhiteSpace(mail.Text) ||
                    mailDataBase.SelectedIndex < 0)
                {
                    MessageBox.Show("Не все обязательные поля заполнены!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                #endregion
                #region Сохранение входных данных для использования в другом потоке
                // Этап 1
                CreateDataStage = 1;
                newUserData.Surname = surname.Text;
                newUserData.Name = name.Text;
                newUserData.PlaceInAD = placeInAD.Text;
                newUserData.Login = login.Text;
                newUserData.NameInAD = nameInAD.Text;
                // Этап 2
                CreateDataStage = 2;
                if (city.SelectedItem != null)
                {
                    newUserData.City = ((Cities)city.SelectedItem).Name;
                    if (((Cities)city.SelectedItem).Groups.Count > 0)
                    {
                        List<string> groupsToAdd = new List<string>();
                        for (int i = 0; i < ((Cities)city.SelectedItem).Groups.Count; i++)
                        {
                            if (!string.IsNullOrWhiteSpace(((Cities)city.SelectedItem).Groups[i]))
                                groupsToAdd.Add(((Cities)city.SelectedItem).Groups[i]);
                        }
                        newUserData.Groups = groupsToAdd;
                    }
                }
                else
                {
                    newUserData.City = "";
                }
                // Этап 3
                CreateDataStage = 3;
                if (organization.SelectedIndex >= 0)
                    newUserData.Organization = organization.SelectedItem.ToString();
                else
                    newUserData.Organization = "";
                // Этап 4
                CreateDataStage = 4;
                newUserData.Adress = adress.Text;
                newUserData.TelephoneNumber = telephoneNumber.Text;
                newUserData.Mobile = mobile.Text;
                newUserData.Post = post.Text;
                newUserData.Department = department.Text;
                newUserData.Mail = mail.Text;
                newUserData.MailDataBase = mailDataBase.SelectedItem.ToString();
                // Этап 5
                CreateDataStage = 5;
                if (chActiveSync.IsChecked == true)
                    newUserData.ActiveSync = true;
                else
                    newUserData.ActiveSync = false;
                // Этап 6
                CreateDataStage = 6;
                if (activeSyncPolicy.SelectedIndex >= 0)
                    newUserData.ActiveSyncPolicy = activeSyncPolicy.SelectedItem.ToString();
                else
                    newUserData.ActiveSyncPolicy = "";
                // Этап 7
                CreateDataStage = 7;
                if (chOWA.IsChecked == true)
                    newUserData.Owa = true;
                else
                    newUserData.Owa = false;
                // Этап 8
                CreateDataStage = 8;
                if (owaPolicy.SelectedIndex >= 0)
                    newUserData.OwaPolicy = owaPolicy.SelectedItem.ToString();
                else
                    newUserData.OwaPolicy = "";
                // Этап 9
                CreateDataStage = 9;
                if (_groups.Count > 0)
                {
                    List<string> groupsForCheck = newUserData.Groups;
                    bool isFinded = false;
                    for (int i = 0; i < _groups.Count; i++)
                    {
                        isFinded = false;
                        for (int j = 0; j < groupsForCheck.Count; j++)
                        {
                            if(_groups[i] == groupsForCheck[j])
                            {
                                isFinded = true;
                                break;
                            }
                        }
                        if(!isFinded)
                            newUserData.Groups.Add(_groups[i]);
                    }
                }
                #endregion
                statusText.Foreground = new SolidColorBrush(Colors.Black);
                statusText.Content = "Создание учетной записи пользователя...";
                IsEnableForm(false);
                _processRun.Clear();
                _processRun.Add(new ProcessRunData { Text = "----Создание учетной записи пользователя----" });
            }
            catch (Exception exp)
            {
                MessageBox.Show("Ошибка на этапе " + CreateDataStage.ToString() + ":\r\n" + exp.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Thread t = new Thread(() =>
            {
                ProcessRunData itemProcess = new ProcessRunData();
                #region Проверка выбранных групп в домене
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusText.Content = "Проверка выбранных групп в домене...";
                    _processRun.Add(new ProcessRunData { Text = "Проверка выбранных групп в домене..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    SearchResult searchResult;
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    for (int i = 0; i < newUserData.Groups.Count; i++)
                    {
                        dirSearcher.Filter = string.Format("(&(objectClass=group)(sAMAccountName={0}))", newUserData.Groups[i]);
                        searchResult = dirSearcher.FindOne();
                        if (searchResult == null)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = "Ошибка проверки выбранных групп в домене:\r\nГруппа " + newUserData.Groups[i] + " отсутствует в домене!!!";
                                _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                                statusText.Foreground = new SolidColorBrush(Colors.Red);
                                statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                                IsEnableForm(true);
                            }));
                            return;
                        }
                    }
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка проверки выбранных групп в домене:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Проверка выбранных групп в домене успешно выполнена";
                }));
                #endregion

                #region Создание учетной записи в АД
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusText.Content = "Создание учетной записи пользователя этап 1/4...";
                    _processRun.Add(new ProcessRunData { Text = "Создание учетной записи в АД..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                int stage = 0;
                try
                {
                    DirectoryEntry ouEntry = null;
                    DirectorySearcher dirSearcher = new DirectorySearcher(_ADSession);
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.Filter = string.Format("(&(|(objectClass=organizationalUnit)(objectClass=organization)(cn=Users)(cn=Computers))(distinguishedName=" + newUserData.PlaceInAD.Replace("\\", "\\5C").Replace("*", "\\2a") + "))");
                    SearchResult searchResult = dirSearcher.FindOne();
                    if (searchResult != null)
                    {
                        // Создание учетной записи
                        stage = 1;
                        ouEntry = searchResult.GetDirectoryEntry();
                        DirectoryEntry childEntry = ouEntry.Children.Add("CN=" + newUserData.NameInAD, "user");
                        childEntry.Properties["sn"].Add(newUserData.Surname);
                        childEntry.Properties["givenName"].Add(newUserData.Name);
                        childEntry.Properties["sAMAccountName"].Add(newUserData.Login);
                        childEntry.Properties["userPrincipalName"].Add(newUserData.Login + "@" + configData["userprincipalname"]);
                        childEntry.Properties["displayName"].Add(newUserData.NameInAD);
                        childEntry.CommitChanges();
                        // Установка остальных значений
                        stage = 2;
                        if (!string.IsNullOrWhiteSpace(newUserData.City))
                            childEntry.Properties["l"].Value = newUserData.City;
                        if (!string.IsNullOrWhiteSpace(newUserData.Organization))
                            childEntry.Properties["company"].Value = newUserData.Organization;
                        if (!string.IsNullOrWhiteSpace(newUserData.Adress))
                            childEntry.Properties["streetAddress"].Value = newUserData.Adress;
                        if (!string.IsNullOrWhiteSpace(newUserData.TelephoneNumber))
                            childEntry.Properties["telephoneNumber"].Value = newUserData.TelephoneNumber;
                        if (!string.IsNullOrWhiteSpace(newUserData.Mobile))
                            childEntry.Properties["mobile"].Value = newUserData.Mobile;
                        if (!string.IsNullOrWhiteSpace(newUserData.Post))
                            childEntry.Properties["title"].Value = newUserData.Post;
                        if (!string.IsNullOrWhiteSpace(newUserData.Department))
                            childEntry.Properties["department"].Value = newUserData.Department;
                        childEntry.CommitChanges();
                        ouEntry.CommitChanges();
                        // Установка стандартного пароля, требование сменить пароль при входе, включение учетной записи
                        stage = 3;
                        using (UserPrincipal oUserPrincipal = UserPrincipal.FindByIdentity(_principalContext, IdentityType.SamAccountName, newUserData.Login))
                        {
                            oUserPrincipal.SetPassword(configData["userdefaultpassword"]);
                            oUserPrincipal.Save();
                            oUserPrincipal.ExpirePasswordNow();
                            oUserPrincipal.Save();
                            oUserPrincipal.Enabled = true;
                            oUserPrincipal.Save();
                        }
                        // Добавление группы
                        stage = 4;
                        newUserData.Groups.ForEach(x =>
                        {
                            dirSearcher.Filter = string.Format("(&(objectClass=group)(sAMAccountName={0}))", x);
                            dirSearcher.PropertiesToLoad.Add("member");
                            searchResult = dirSearcher.FindOne();
                            if (searchResult != null)
                            {
                                DirectoryEntry group = searchResult.GetDirectoryEntry();
                                group.Properties["member"].Add("CN=" + newUserData.NameInAD + "," + newUserData.PlaceInAD);
                                group.CommitChanges();
                            }
                        });
                        Thread.Sleep(15000);
                    }
                    else
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                            itemProcess.Text = "Ошибка создания учетной записи в АД, этап " + stage.ToString() + ":\r\n" + "Не найдено место расположения в АД " + newUserData.PlaceInAD;
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            statusText.Foreground = new SolidColorBrush(Colors.Red);
                            statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка создания учетной записи в АД, этап " + stage.ToString() + ":\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Учетная запись в АД успешно создана";
                }));
                #endregion

                #region Создание учетной записи в Exchange
                // Настройка параметров подключения к Exchange server
                //============================================================
                PSCredential credential = new PSCredential(_login, _password);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo((new Uri(configData["mailserver"])), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
                Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(connectionInfo);
                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                string[] arrMailDomains = configData["maildomain"].Trim().Split(',');
                string mailUserName = newUserData.Mail.Split('@')[0];
                string[] mailAddressToAdd = new string[] { };
                //============================================================

                // Создание массива с эл. адресами
                for (int i = 0; i < arrMailDomains.Length; i++)
                {
                    if (arrMailDomains[i] == configData["mailprimarydomain"].Trim())
                    {
                        Array.Resize(ref mailAddressToAdd, mailAddressToAdd.Length + 1);
                        mailAddressToAdd[mailAddressToAdd.Length - 1] = "SMTP:" + mailUserName + "@" + arrMailDomains[i];
                    }
                    else
                    {
                        Array.Resize(ref mailAddressToAdd, mailAddressToAdd.Length + 1);
                        mailAddressToAdd[mailAddressToAdd.Length - 1] = "smtp:" + mailUserName + "@" + arrMailDomains[i];
                    }
                }

                #region Подключение к Exchange Server
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusText.Content = "Создание учетной записи пользователя этап 2/4...";
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
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка подключения к Exchange Server:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Подключение к Exchange Server выполнено успешно";
                }));
                #endregion

                #region Создание учетной записи в Exchange
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusText.Content = "Создание учетной записи пользователя этап 3/4...";
                    _processRun.Add(new ProcessRunData { Text = "Создание учетной записи в Exchange..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    
                    #region Создание учетной записи в Exchange
                    stage = 1;
                    command.AddCommand("Enable-Mailbox");
                    command.AddParameter("Identity", newUserData.Login);
                    command.AddParameter("Alias", newUserData.Login);
                    command.AddParameter("Database", newUserData.MailDataBase);
                    powershell.Commands = command;
                    powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка cоздания учетной записи в Exchange:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = errorMsg;
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            statusText.Foreground = new SolidColorBrush(Colors.Red);
                            statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                    #endregion
                    #region Установка электронных адресов
                    stage = 2;
                    command.AddCommand("Set-Mailbox");
                    command.AddParameter("Identity", newUserData.Login);
                    command.AddParameter("EmailAddressPolicyEnabled", false);
                    command.AddParameter("EmailAddresses", mailAddressToAdd);
                    powershell.Commands = command;
                    powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка установки эл.адресов:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = errorMsg;
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            statusText.Foreground = new SolidColorBrush(Colors.Red);
                            statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                    #endregion
                    #region Настройка ActiveSync и OWA
                    stage = 3;
                    command.AddCommand("Set-CASMailbox");
                    command.AddParameter("Identity", newUserData.Login);
                    //command.AddParameter("ImapEnabled", false);
                    //command.AddParameter("PopEnabled", false);
                    command.AddParameter("ActiveSyncEnabled", newUserData.ActiveSync);
                    command.AddParameter("ActiveSyncMailboxPolicy", newUserData.ActiveSyncPolicy);
                    command.AddParameter("OWAEnabled", newUserData.Owa);
                    command.AddParameter("OWAMailboxPolicy", newUserData.OwaPolicy);
                    powershell.Commands = command;
                    powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка настройки ActiveSync и OWA:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text = errorMsg;
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                            statusText.Foreground = new SolidColorBrush(Colors.Red);
                            statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                    #endregion
                    Thread.Sleep(15000);
                    #region Руссификация папок
                    //stage = 4;
                    //command.AddCommand("Set-MailboxRegionalConfiguration");
                    //command.AddParameter("Identity", newUserData.Login);
                    //command.AddParameter("Language", "ru-RU");
                    //command.AddParameter("LocalizeDefaultFolderName", true);
                    //powershell.Commands = command;
                    //powershell.Invoke();
                    //command.Clear();
                    //if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    //{
                    //    for (int i = 0; i < powershell.Streams.Error.Count; i++)
                    //    {
                    //        string errorMsg = "Ошибка руссификации папок:\r\n" + powershell.Streams.Error[i].ToString();
                    //        Dispatcher.BeginInvoke(new Action(() =>
                    //        {
                    //            itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                    //            itemProcess.Text = errorMsg;
                    //        }));
                    //    }
                    //}
                    #endregion
                }
                catch (Exception exp)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка cоздания учетной записи в Exchange этап " + stage.ToString() + ":\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе создания учетной записи возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (itemProcess.Image != null)
                    {
                        itemProcess.Text += "\r\nСоздание учетной записи в Exchange завершено";
                    }
                    else
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Учетная запись в Exchange успешно создана";
                    }
                }));
                #endregion
                // Закрытие соединений с Exchange server
                //============================================================
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
                command.Clear();
                //============================================================
                #endregion

                #region Очистка введенных данных
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    statusText.Content = "Создание учетной записи пользователя этап 4/4...";
                    if (chCopySettings.IsChecked == true)
                    {
                        surname.Text = "";
                        name.Text = "";
                        middlename.Text = "";
                        login.Text = "";
                        nameInAD.Text = "";
                        mail.Text = "";
                    }
                    else
                    {
                        surname.Text = "";
                        name.Text = "";
                        middlename.Text = "";
                        login.Text = "";
                        placeInAD.Text = "";
                        nameInAD.Text = "";
                        mail.Text = "";
                        city.SelectedIndex = -1;
                        city.Text = "";
                        organization.SelectedIndex = Int32.Parse(configData["indexdefaultorganizations"]) - 1;
                        adress.Text = "";
                        telephoneNumber.Text = "";
                        mobile.Text = "";
                        post.Text = "";
                        post.SelectedIndex = -1;
                        department.Text = "";
                        department.SelectedIndex = -1;
                        mail.Text = "";
                        mailDataBase.SelectedIndex = -1;
                        chActiveSync.IsChecked = false;
                        labelActiveSyncPolicy.IsEnabled = false;
                        activeSyncPolicy.IsEnabled = false;
                        chOWA.IsChecked = false;
                        labelOwaPolicy.IsEnabled = false;
                        owaPolicy.IsEnabled = false;
                        templates.SelectedIndex = -1;
                        _groups.Clear();
                    }
                    _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    statusText.Foreground = new SolidColorBrush(Colors.Green);
                    statusText.Content = "Учетная запись " + newUserData.NameInAD + " успешно создана.";
                    IsEnableForm(true);
                }));
                #endregion
            });
            t.IsBackground = true;
            t.Start();
        }
        #endregion

        #region Методы
        // Чтение конфигурационного файла
        private Dictionary<string, string> GetConfigData(string configFilePach, ref string errorStr)
        {
            Dictionary<string, string> configData = new Dictionary<string, string>();
            string configDataFull = "";

            try
            {
                using (StreamReader sr = new StreamReader(configFilePach, Encoding.GetEncoding("windows-1251")))
                {
                    configDataFull = sr.ReadToEnd();
                    string[] configArr = configDataFull.Split('\n');
                    for (int i = 0; i < configArr.Length; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(configArr[i]) && !configArr[i].StartsWith("#"))
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
        // Транслитерация
        private string Translit(string inputText)
        {
            Dictionary<char, string> translitDictionary = new Dictionary<char, string>() 
            { 
                {'а', "a"}, {'А', "A"}, {'б', "b"}, {'Б', "B"},
                {'в', "v"}, {'В', "V"}, {'г', "g"}, {'Г', "G"},
                {'д', "d"}, {'Д', "D"}, {'е', "e"}, {'Е', "E"},
                {'ё', "e"}, {'Ё', "E"}, {'ж', "zh"}, {'Ж', "Zh"},
                {'з', "z"}, {'З', "Z"}, {'и', "i"}, {'И', "I"},
                {'й', "y"}, {'Й', "Y"}, {'к', "k"}, {'К', "K"},
                {'л', "l"}, {'Л', "L"}, {'м', "m"}, {'М', "M"},
                {'н', "n"}, {'Н', "N"}, {'о', "o"}, {'О', "O"},
                {'п', "p"}, {'П', "P"}, {'р', "r"}, {'Р', "R"},
                {'с', "s"}, {'С', "S"}, {'т', "t"}, {'Т', "T"},
                {'у', "u"}, {'У', "U"}, {'ф', "f"}, {'Ф', "F"},
                {'х', "kh"}, {'Х', "Kh"}, {'ц', "ts"}, {'Ц', "Ts"},
                {'ч', "ch"}, {'Ч', "Ch"}, {'ш', "sh"}, {'Ш', "Sh"},
                {'щ', "shch"}, {'Щ', "Shch"}, {'ъ', ""}, {'Ъ', ""},
                {'ы', "y"}, {'Ы', "Y"}, {'ь', ""}, {'Ь', ""},
                {'э', "e"}, {'Э', "E"}, {'ю', "yu"}, {'Ю', "Yu"},
                {'я', "ya"}, {'Я', "Ya"}
            }; // Словарь сопоставлений русских и английских букв
            char[] inChars = inputText.ToCharArray();
            string result = "";
            foreach (char c in inChars)
            {
                if (translitDictionary.ContainsKey(c))
                {
                    result += translitDictionary[c];
                }
                else
                {
                    result += c;
                }
            }
            return result;
        }
        // Создание логина
        private string CreateLogin()
        {
            string login = "";
            if (!string.IsNullOrWhiteSpace(name.Text) && !string.IsNullOrWhiteSpace(surname.Text) && !string.IsNullOrWhiteSpace(middlename.Text))
            {
                string nameTl = Translit(name.Text.Substring(0, 1));
                string middlenameTl = Translit(middlename.Text.Substring(0, 1));
                login = Translit(surname.Text + "." + nameTl.Substring(0,1) + middlenameTl.Substring(0,1));
            }
            return login;
        }
        // Создание имени в АД
        private string CreateNameInAD()
        {
            string nameInAd = "";
            if (!string.IsNullOrWhiteSpace(name.Text) && !string.IsNullOrWhiteSpace(surname.Text) && !string.IsNullOrWhiteSpace(middlename.Text))
            {
                nameInAd = surname.Text + " " + name.Text + " " + middlename.Text;
            }
            if (nameInAd.Length > 64)
                return nameInAd.Substring(0, 64);
            else
                return nameInAd;
        }
        // Создание электронного адреса
        private string CreateMail()
        {
            string mailAdress = "";
            if (!string.IsNullOrWhiteSpace(name.Text) && !string.IsNullOrWhiteSpace(surname.Text) && !string.IsNullOrWhiteSpace(middlename.Text) && configData.ContainsKey("mailprimarydomain") && !string.IsNullOrWhiteSpace(configData["mailprimarydomain"]))
            {
                string nameToMail = Translit(name.Text.Substring(0, 1));
                string surnameToMail = Translit(surname.Text);
                string middlenameToMail = Translit(middlename.Text.Substring(0, 1));
                mailAdress = surnameToMail + "." + nameToMail.Substring(0, 1) + middlenameToMail.Substring(0,1) + "@" + configData["mailprimarydomain"];
            }
            return mailAdress;
        }
        // Загрузка данных конфигурационного файла и необходимой информации из Exchange
        private void LoadConfigData()
        {
            statusText.Content = "Загрузка данных из конфигурационных файлов и Exchange...";
            IsEnableForm(false);
            _processRun.Clear();
            string errorMesg = "";
            string dataFull = "";
            _cities = new ObservableCollection<Cities>();
            _posts = new ObservableCollection<string>();
            _departments = new ObservableCollection<string>();
            _mailDataBase = new ObservableCollection<string>();
            _activeSyncPolicy = new ObservableCollection<string>();
            _owaPolicy = new ObservableCollection<string>();
            _templates = new ObservableCollection<string>();
            _processRun.Add(new ProcessRunData { Text = "----Загрузка данных из конфигурационных файлов и Exchange----" });
            Thread t = new Thread(() =>
            {
                ProcessRunData itemProcess = new ProcessRunData();
                
                #region Чтение конфигурационного файла плагина
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
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка чтения конфигурационного файла:\r\n" + errorMesg;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                if (!configData.ContainsKey("mailserver") || 
                    !configData.ContainsKey("maildomain") || 
                    !configData.ContainsKey("mailprimarydomain") || 
                    !configData.ContainsKey("organizations") ||
                    !configData.ContainsKey("indexdefaultorganizations") ||
                    !configData.ContainsKey("userprincipalname") ||
                    !configData.ContainsKey("userdefaultpassword"))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Отсутсвуют требуемые параметры в конфигурационном файле!!!";
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                else if (string.IsNullOrWhiteSpace(configData["mailserver"]) ||
                    string.IsNullOrWhiteSpace(configData["maildomain"]) ||
                    string.IsNullOrWhiteSpace(configData["mailprimarydomain"]) ||
                    string.IsNullOrWhiteSpace(configData["organizations"]) ||
                    string.IsNullOrWhiteSpace(configData["indexdefaultorganizations"]) ||
                    string.IsNullOrWhiteSpace(configData["userprincipalname"]) ||
                    string.IsNullOrWhiteSpace(configData["userdefaultpassword"]))
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Один из параметров в конфигурационном файле не имеет значения!!!";
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Чтение конфигурационного файла плагина выполнено";
                }));
                #endregion

                #region Загрузка списка городов для выбора
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Загрузка списка городов..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (File.Exists(_pathToPluginDir + "\\cities.dat"))
                    {
                        using (StreamReader sr = new StreamReader(_pathToPluginDir + "\\cities.dat", Encoding.GetEncoding("windows-1251")))
                        {
                            dataFull = sr.ReadToEnd();
                            string[] configArr = dataFull.Split('\n');
                            for (int i = 0; i < configArr.Length; i++)
                            {
                                if (!string.IsNullOrWhiteSpace(configArr[i]))
                                {
                                    List<string> cityGroups = new List<string>();
                                    string[] city = configArr[i].TrimEnd('\r').Split(';');
                                    string[] cityGroup = city[3].Split(',');
                                    if (cityGroup.Length > 0)
                                        for (int j = 0; j < cityGroup.Length; j++)
                                        {
                                            if (!string.IsNullOrWhiteSpace(cityGroup[j]))
                                                cityGroups.Add(cityGroup[j]);
                                        }
                                    _cities.Add(new Cities { DisplayName = city[0], Name = city[1], Adress = city[2], Groups = cityGroups });
                                }
                            }
                        }
                    }
                    //_cities = new ObservableCollection<Cities>(_cities.OrderBy(i => i.DisplayName));
                }
                catch (Exception e)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка загрузки списка городов:\r\n" + e.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Загрузка списка городов завершена";
                }));
                #endregion

                #region Загрузка списка должностей для выбора
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Загрузка списка должностей..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (File.Exists(_pathToPluginDir + "\\posts.dat"))
                    {
                        using (StreamReader sr = new StreamReader(_pathToPluginDir + "\\posts.dat", Encoding.GetEncoding("windows-1251")))
                        {
                            dataFull = sr.ReadToEnd();
                            string[] configArr = dataFull.Split('\n');
                            for (int i = 0; i < configArr.Length; i++)
                            {
                                if (!string.IsNullOrWhiteSpace(configArr[i]))
                                {
                                    _posts.Add(configArr[i].TrimEnd('\r'));
                                }
                            }
                        }
                    }
                    _posts = new ObservableCollection<string>(_posts.OrderBy(i => i));
                }
                catch (Exception e)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка загрузки списка должностей:\r\n" + e.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Загрузка списка должностей завершена";
                }));
                #endregion

                #region Загрузка списка отделов для выбора
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Загрузка списка отделов..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (File.Exists(_pathToPluginDir + "\\departments.dat"))
                    {
                        using (StreamReader sr = new StreamReader(_pathToPluginDir + "\\departments.dat", Encoding.GetEncoding("windows-1251")))
                        {
                            dataFull = sr.ReadToEnd();
                            string[] configArr = dataFull.Split('\n');
                            for (int i = 0; i < configArr.Length; i++)
                            {
                                if (!string.IsNullOrWhiteSpace(configArr[i]))
                                {
                                    _departments.Add(configArr[i].TrimEnd('\r'));
                                }
                            }
                        }
                    }
                    _departments = new ObservableCollection<string>(_departments.OrderBy(i => i));
                }
                catch (Exception e)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка загрузки списка отделов:\r\n" + e.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Загрузка списка отделов завершена";
                }));
                #endregion

                // Настройка параметров подключения к Exchange server
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
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка подключения к Exchange Server:\r\n" + exp.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Подключение к Exchange Server выполнено успешно";
                }));
                #endregion

                #region Загрузка списка почтовых баз данных из Exchange
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Загрузка списка почтовых баз данных из Exchange..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    command.AddCommand("Get-MailboxDatabase");
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", "Name");
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Text = "";
                        }));
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки списка почтовых баз данных из Exchange:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text += errorMsg;
                                _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                                
                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusText.Foreground = new SolidColorBrush(Colors.Red);
                            statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        for (int i = 0; i < result.Count; i++)
                        {
                            _mailDataBase.Add(result[i].Properties["Name"].Value.ToString());
                        }
                        // сортировка полученных баз даных в алфавитном порядке
                        _mailDataBase = new ObservableCollection<string>(_mailDataBase.OrderBy(i => i));
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка загрузки списка почтовых баз данных из Exchange:\r\n" + ex.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    }));
                    command.Clear();
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Загрузка списка почтовых баз данных из Exchange выполнена успешно";
                }));
                #endregion

                #region Загрузка списка политик ActiveSync из Exchange
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Загрузка списка политик ActiveSync из Exchange..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    command.AddCommand("Get-ActiveSyncMailboxPolicy");
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", "Identity");
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Text = "";
                        }));
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки списка политик ActiveSync из Exchange:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text += errorMsg;
                                _processRun.Add(new ProcessRunData { Text = "----Завершено----" });

                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusText.Foreground = new SolidColorBrush(Colors.Red);
                            statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        for (int i = 0; i < result.Count; i++)
                        {
                            _activeSyncPolicy.Add(result[i].Properties["Identity"].Value.ToString());
                        }
                        // сортировка полученных политик в алфавитном порядке
                        _activeSyncPolicy = new ObservableCollection<string>(_activeSyncPolicy.OrderBy(i => i));
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка загрузки списка политик ActiveSync из Exchange:\r\n" + ex.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    }));
                    command.Clear();
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Загрузка списка политик ActiveSync из Exchange выполнена успешно";
                }));
                #endregion

                #region Загрузка списка политик OWA из Exchange
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Загрузка списка политик OWA из Exchange..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    command.AddCommand("Get-OwaMailboxPolicy");
                    command.AddCommand("Select-Object");
                    command.AddParameter("Property", "Identity");
                    powershell.Commands = command;
                    var result = powershell.Invoke();
                    command.Clear();
                    if (powershell.Streams.Error != null && powershell.Streams.Error.Count > 0)
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            itemProcess.Text = "";
                        }));
                        for (int i = 0; i < powershell.Streams.Error.Count; i++)
                        {
                            string errorMsg = "Ошибка загрузки списка политик OWA из Exchange:\r\n" + powershell.Streams.Error[i].ToString();
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                                itemProcess.Text += errorMsg;
                                _processRun.Add(new ProcessRunData { Text = "----Завершено----" });

                            }));
                        }
                        runspace.Dispose();
                        runspace = null;
                        powershell.Dispose();
                        powershell = null;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            statusText.Foreground = new SolidColorBrush(Colors.Red);
                            statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                            IsEnableForm(true);
                        }));
                        return;
                    }
                    if (result.Count > 0)
                    {
                        for (int i = 0; i < result.Count; i++)
                        {
                            _owaPolicy.Add(result[i].Properties["Identity"].Value.ToString());
                        }
                        // сортировка полученных политик в алфавитном порядке
                        _owaPolicy = new ObservableCollection<string>(_owaPolicy.OrderBy(i => i));
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка загрузки списка политик OWA из Exchange:\r\n" + ex.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                    }));
                    command.Clear();
                    runspace.Dispose();
                    runspace = null;
                    powershell.Dispose();
                    powershell = null;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Загрузка списка политик OWA из Exchange выполнена успешно";
                }));
                #endregion

                // Закрытие соединений с Exchange server
                //============================================================
                runspace.Dispose();
                runspace = null;
                powershell.Dispose();
                powershell = null;
                command.Clear();
                //============================================================

                #region Формирование списка шаблонов
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Формирование списка шаблонов..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                }));
                try
                {
                    if (Directory.Exists(_pathToPluginDir + "\\templates"))
                    {
                        string[] arrTemplates = Directory.GetFiles(_pathToPluginDir + "\\templates");
                        if (arrTemplates.Length > 0)
                        {
                            for (int j = 0; j < arrTemplates.Length; j++)
                            {
                                _templates.Add(arrTemplates[j].Substring(arrTemplates[j].LastIndexOf("\\")).TrimStart('\\').TrimEnd(new char[] { '.', 't', 'x', 't' }));
                            }
                        }
                        _templates = new ObservableCollection<string>(_templates.OrderBy(i => i));
                    }
                }
                catch (Exception e)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка формирования списка шаблонов:\r\n" + e.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }));
                    return;
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                    itemProcess.Text = "Формирование списка шаблонов завершено";
                }));
                #endregion

                #region Установка загруженных значений на форму
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _processRun.Add(new ProcessRunData { Text = "Установка загруженных значений на форму..." });
                    itemProcess = _processRun[_processRun.Count - 1];
                    try
                    {
                        city.ItemsSource = _cities;
                        city.DisplayMemberPath = "DisplayName";

                        organization.Items.Clear();
                        string[] arrOrganizations = configData["organizations"].Split(',');
                        for (int i = 0; i < arrOrganizations.Length; i++)
                        {
                            organization.Items.Add(arrOrganizations[i]);
                        }
                        if (Int32.Parse(configData["indexdefaultorganizations"]) > 0)
                            organization.SelectedIndex = Int32.Parse(configData["indexdefaultorganizations"])-1;
                        post.ItemsSource = _posts;
                        department.ItemsSource = _departments;
                        mailDataBase.ItemsSource = _mailDataBase;
                        activeSyncPolicy.ItemsSource = _activeSyncPolicy;
                        if (_activeSyncPolicy.Count > 0)
                            activeSyncPolicy.SelectedIndex = 0;
                        owaPolicy.ItemsSource = _owaPolicy;
                        if (_owaPolicy.Count > 0)
                            owaPolicy.SelectedIndex = 0;
                        templates.ItemsSource = _templates;


                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/select.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Установка загруженных значений на форму завершена";
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Content = "";
                        IsEnableForm(true);
                    }
                    catch (Exception e)
                    {
                        itemProcess.Image = new BitmapImage(new Uri(@"/CreateUsersPPK;component/Resources/stop.ico", UriKind.RelativeOrAbsolute));
                        itemProcess.Text = "Ошибка установки загруженных значений на форму:\r\n" + e.Message;
                        _processRun.Add(new ProcessRunData { Text = "----Завершено----" });
                        statusText.Foreground = new SolidColorBrush(Colors.Red);
                        statusText.Content = "В процессе загрузки возникли проблемы, детальную информацию смотрите в отчете.";
                        IsEnableForm(true);
                    }
                }));
                #endregion
            });
            t.IsBackground = true;
            t.Start();
        }
        // Отключение и включение формы
        private void IsEnableForm(bool enable)
        {
            surname.IsEnabled = enable;
            name.IsEnabled = enable;
            middlename.IsEnabled = enable;
            placeInAD.IsEnabled = enable;
            btSelectOU.IsEnabled = enable;
            login.IsEnabled = enable;
            checkLogin.IsEnabled = enable;
            nameInAD.IsEnabled = enable;
            checkNameInAD.IsEnabled = enable;
            city.IsEnabled = enable;
            btAddCity.IsEnabled = enable;
            btEditCity.IsEnabled = enable;
            organization.IsEnabled = enable;
            adress.IsEnabled = enable;
            telephoneNumber.IsEnabled = enable;
            mobile.IsEnabled = enable;
            post.IsEnabled = enable;
            department.IsEnabled = enable;
            mail.IsEnabled = enable;
            btCheckMail.IsEnabled = enable;
            mailDataBase.IsEnabled = enable;
            chActiveSync.IsEnabled = enable;
            chOWA.IsEnabled = enable;
            if (enable && chActiveSync.IsChecked == true)
            {
                activeSyncPolicy.IsEnabled = true;
            }
            else if (enable && chActiveSync.IsChecked == false)
            {
                activeSyncPolicy.IsEnabled = false;
            }
            else
            {
                activeSyncPolicy.IsEnabled = false;
            }
            if (enable && chOWA.IsChecked == true)
            {
                owaPolicy.IsEnabled = true;
            }
            else if (enable && chOWA.IsChecked == false)
            {
                owaPolicy.IsEnabled = false;
            }
            else
            {
                owaPolicy.IsEnabled = false;
            }
            templates.IsEnabled = enable;
            btAddGroups.IsEnabled = enable;
            chCopySettings.IsEnabled = enable;
            btCreateUser.IsEnabled = enable;
            btReport.IsEnabled = enable;
            btReloadConfig.IsEnabled = enable;
            if(enable && templates.SelectedIndex >= 0)
            {
                btLoadTemplate.IsEnabled = true;
            }
            else if (enable && templates.SelectedIndex < 0)
            {
                btLoadTemplate.IsEnabled = false;
            }
            else
            {
                btLoadTemplate.IsEnabled = false;
            }
            btAddTemplate.IsEnabled = enable;
            btEditTemplate.IsEnabled = enable;
            btUpdateTemplate.IsEnabled = enable;
            if(!enable)
                btDeleteGroups.IsEnabled = enable;
            else
            {
                Binding binding = new Binding();
                binding.Source = goupListToAdd; // установить в качестве source object значение ElementName
                binding.Path = new PropertyPath(ListView.SelectedItemProperty);
                binding.Mode = BindingMode.OneWay;
                binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
                binding.Converter = new NullToBooleanConverter();
                BindingOperations.SetBinding(btDeleteGroups, Button.IsEnabledProperty, binding);
            }
            btSavePost.IsEnabled = enable;
            btSaveDepartment.IsEnabled = enable;
        }
        #endregion
    }
}
