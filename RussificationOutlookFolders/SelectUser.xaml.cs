using RussificationOutlookFolders.DataProvider;
using RussificationOutlookFolders.Model;
using RussificationOutlookFolders.UISearchTextBox;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;

namespace RussificationOutlookFolders
{
    /// <summary>
    /// Логика взаимодействия для SelectUser.xaml
    /// </summary>
    public partial class SelectUser : Window
    {
        private Dictionary<string, string> fieldsInAD; // Сопоставление групп поиска и полей в АД
        private PrincipalContext _principalContext; // Контекст соединения с АД
        private DirectoryEntry _sessionAD; //сессия АД

        public string SelectedUser { get; set; }

        // Конструктор
        public SelectUser(DirectoryEntry entry, PrincipalContext context, string findStr = "")
        {
            InitializeComponent();
            _principalContext = context;
            _sessionAD = entry;
            #region Сопоставление групп поиска и полей в АД
            fieldsInAD = new Dictionary<string, string>();
            fieldsInAD.Add("По умолчанию", "Default");
            fieldsInAD.Add("Имя пользователя в АД", "name");
            fieldsInAD.Add("Отображаемое имя", "displayName");
            fieldsInAD.Add("Организация", "company");
            fieldsInAD.Add("Имя", "givenName");
            fieldsInAD.Add("Фамилия", "sn");
            fieldsInAD.Add("Логин", "sAMAccountName");
            fieldsInAD.Add("Должность", "title");
            fieldsInAD.Add("Отдел", "department");
            fieldsInAD.Add("Адрес", "streetAddress");
            fieldsInAD.Add("Город", "l");
            fieldsInAD.Add("Эл. Почта", "mail");
            fieldsInAD.Add("Телефон внутренний", "telephoneNumber");
            fieldsInAD.Add("Телефон мобильный", "mobile");
            #endregion
            // Устанавливаем настройки поиска пользователя в домене, создаём группы поиска и выбираем стиль отображения групп
            List<string> sections = new List<string> { 
                "По умолчанию", 
                "Имя пользователя в АД", 
                "Отображаемое имя", 
                "Организация",
                "Имя", 
                "Фамилия", 
                "Логин", 
                "Должность", 
                "Отдел", 
                "Адрес",
                "Город", 
                "Эл. Почта", 
                "Телефон внутренний", 
                "Телефон мобильный" };
            Search.SectionsList = sections;
            Search.SectionsStyle = SearchTextBox.SectionsStyles.RadioBoxStyle;
            // Определяем обработчик поиска
            Search.OnSearch += new RoutedEventHandler(search_OnSearch);
            Search.Focus();

            if (!string.IsNullOrWhiteSpace(findStr))
            {
                Search.Text = findStr;
                search_OnSearch("defaulData", new RoutedEventArgs());
            }
        }

        // Запущен процесс поиска пользователя в АД
        private void search_OnSearch(object sender, RoutedEventArgs e)
        {
            SearchEventArgs searchArgs = null;
            string selectedSection = "";
            string filterStr = Search.Text;
            if (sender is string)
            {
                if ((string)sender == "defaulData")
                {
                    selectedSection = Search.SectionsList[0];
                }
            }
            else
                searchArgs = e as SearchEventArgs;
            ReadOnlyCollection<User> items;
            string errorMsg = "";

            ListUsersForSelected.ItemsSource = null;
            Search.IsEnabled = false;
            Filter.IsEnabled = false;
            Filter.Text = "";
            Thread t = new Thread(() =>
            {
                if (searchArgs != null)
                    items = AsyncDataProvider.GetUsersForSelected(fieldsInAD[searchArgs.Sections[0]], searchArgs.Keyword, _principalContext, _sessionAD, ref errorMsg);
                else
                    items = AsyncDataProvider.GetUsersForSelected(fieldsInAD[selectedSection], filterStr, _principalContext, _sessionAD, ref errorMsg);

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    ListUsersForSelected.ItemsSource = items;
                    CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(ListUsersForSelected.ItemsSource);
                    view.Filter = FindedUsers_Filter;
                    if (view.Count > 0)
                    {
                        ListUsersForSelected.SelectedIndex = 0;
                    }
                    Filter.IsEnabled = true;
                    Search.IsEnabled = true;
                    if (!string.IsNullOrWhiteSpace(errorMsg))
                        MessageBox.Show(errorMsg, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    Search.Focus();
                }));
            });
            t.IsBackground = true;
            t.Start();
        }
        // Фильтр найденных пользователей
        private bool FindedUsers_Filter(object item)
        {
            if (String.IsNullOrEmpty(Filter.Text))
                return true;

            var user = (User)item;

            return (user.NameInAD.ToUpper().Contains(Filter.Text.ToUpper()));
        }
        // Нажата кнопка выбора пользователя
        private void btSelectUser_Click(object sender, RoutedEventArgs e)
        {
            SelectedUser = ((User)ListUsersForSelected.SelectedItem).Login;
            DialogResult = true;
            Close();
        }
        // Нажата кнопка в основном окне
        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
        // Выполненно двойное нажатие левой кнопки мышки на елементе списка групп для выбора
        private void ListUsersForSelectedItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            btSelectUser_Click(sender, e);
        }
        // Событие изменения текста в фильтре найденных пользователей
        private void Filter_TextChanged(object sender, TextChangedEventArgs e)
        {
            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(ListUsersForSelected.ItemsSource);
            if (view != null)
            {
                view.Refresh();
                if (view.Count > 0)
                {
                    ListUsersForSelected.SelectedIndex = 0;
                }
            }
        }
    }
}
