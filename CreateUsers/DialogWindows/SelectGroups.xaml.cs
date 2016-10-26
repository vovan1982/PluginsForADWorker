using CreateUsers.DataProvider;
using CreateUsers.Model;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.DirectoryServices;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace CreateUsers.DialogWindows
{
    /// <summary>
    /// Логика взаимодействия для SelectGroups.xaml
    /// </summary>
    public partial class SelectGroups : Window
    {
        #region Поля
        private DirectoryEntry _sessionAD; // Сессия АД для выполнения запросов
        private ObservableCollection<Group> _selectedGroups; // Список выбранных групп для добавления
        #endregion

        #region Свойства
        // Список выбранных групп для добавления
        public ObservableCollection<Group> SelectedGroups
        {
            get { return _selectedGroups; }
            private set { _selectedGroups = value; }
        }
        #endregion

        // Конструктор
        public SelectGroups(DirectoryEntry entry)
        {
            InitializeComponent();
            _selectedGroups = new ObservableCollection<Group>();
            _sessionAD = entry;
            selectedGroups.ItemsSource = _selectedGroups;
            ReadOnlyCollection<Group> items;
            groupsForSelected.ItemsSource = null;
            string errorMsg = "";
            Thread t = new Thread(() =>
            {
                items = AsyncDataProvider.GetGroupForSelected(_sessionAD, ref errorMsg);
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    groupsForSelected.ItemsSource = items;
                    CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(groupsForSelected.ItemsSource);
                    if (view != null)
                    {
                        view.Filter = Groups_Filter;
                        view.SortDescriptions.Clear();
                        view.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
                    }
                    if (!string.IsNullOrWhiteSpace(errorMsg))
                        MessageBox.Show(errorMsg, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }));
            });
            t.IsBackground = true;
            t.Start();
        }

        // Фильтр групп
        private bool Groups_Filter(object item)
        {
            if (String.IsNullOrEmpty(filterGroupsForSelected.Text))
                return true;

            var group = (Group)item;

            return (group.Name.ToUpper().Contains(filterGroupsForSelected.Text.ToUpper()));
        }
        // Изменено содержимое поля фильтрации групп
        private void filterGroupsForSelected_TextChanged(object sender, TextChangedEventArgs e)
        {
            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(groupsForSelected.ItemsSource);
            if (view != null)
            {
                view.Refresh();
                if (view.Count > 0)
                {
                    groupsForSelected.SelectedIndex = 0;
                }
                view.SortDescriptions.Clear();
                view.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
            }
        }
        // Нажата кнопка отмены
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        // Нажата кнопка добавления выбранной группы к списку выбранных
        private void btAddSelectedGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (var item in groupsForSelected.SelectedItems)
                {
                    bool isTake = false;
                    _selectedGroups.ToList().ForEach(x =>
                    {
                        if (x.Name == ((Group)item).Name) isTake = true;
                    });

                    if (!isTake)
                    {
                        _selectedGroups.Add(new Group { Name = ((Group)item).Name, PlaceInAD = ((Group)item).PlaceInAD });
                        selectedGroups.SelectedIndex = 0;
                        btSelectedGroups.IsEnabled = true;
                    }
                }
                groupsForSelected.SelectedItems.Clear();
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Нажата кнопка удаления выбранной группы из списка выбранных групп
        private void btDeleteSelectedGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                while (selectedGroups.SelectedItems.Count > 0)
                {
                    _selectedGroups.Remove((Group)selectedGroups.SelectedItems[0]);
                }
                if (selectedGroups.Items.Count <= 0)
                {
                    btSelectedGroups.IsEnabled = false;
                }
                else
                {
                    selectedGroups.SelectedIndex = 0;
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Выполненно двойное нажатие левой кнопки мышки на елементе списка групп для выбора
        private void groupsForSelectedItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            btAddSelectedGroups_Click(sender, e);
        }
        // Выполненно двойное нажатие левой кнопки мышки на елементе списка групп для добавления
        private void selectedGroupsItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            btDeleteSelectedGroups_Click(sender, e);
        }
        // Нажата кнопка добавления пользователя в выбранные группы
        private void btSelectedGroups_Click(object sender, RoutedEventArgs e)
        {
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
    }
}
