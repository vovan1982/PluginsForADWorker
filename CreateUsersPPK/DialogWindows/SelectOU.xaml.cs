using CreateUsersPPK.Converters;
using CreateUsersPPK.DataProvider;
using CreateUsersPPK.Model;
using System;
using System.Collections.ObjectModel;
using System.DirectoryServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace CreateUsersPPK.DialogWindows
{
    /// <summary>
    /// Логика взаимодействия для SelectOU.xaml
    /// </summary>
    public partial class SelectOU : Window
    {
        #region Поля
        private DirectoryEntry _sessionAD; //сессия АД
        #endregion

        #region Свойства
        // Возвращает DN зачение выбранной OU
        public string SelectedOU
        {
            get { return ((DomainTreeItem)DomainOUTreeView.SelectedItem).Description; }
        }
        #endregion

        //Конструктор
        public SelectOU(DirectoryEntry entry)
        {
            InitializeComponent();
            btSelect.IsEnabled = false;
            _sessionAD = entry;
            CreateDomainOUTree();
        }

        #region События
        // Нажата кнопка отмены
        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        // Нажата кнопка переместить
        private void btSelect_Click(object sender, RoutedEventArgs e)
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
        // Нажата клавиша Enter на элементе дерева
        private void DomainOUTreeViewItem_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter && DomainOUTreeView.SelectedItem != null && !((DomainTreeItem)DomainOUTreeView.SelectedItem).Description.StartsWith("DC"))
            {
                e.Handled = true;
                btSelect_Click(sender, new RoutedEventArgs());
            }
        }
        #endregion

        #region Методы
        // Создание дерева подразделений домена
        private void CreateDomainOUTree()
        {
            ReadOnlyCollection<DomainTreeItem> items;
            DomainOUTreeView.ItemsSource = null;
            string errorMsg = "";
            new Thread(() =>
            {
                items = AsyncDataProvider.GetDomainOUTree(_sessionAD, ref errorMsg);
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    DomainOUTreeView.ItemsSource = items;

                    Binding binding = new Binding();
                    binding.Source = DomainOUTreeView; // установить в качестве source object значение ElementName
                    binding.Path = new PropertyPath("SelectedItem.Description");
                    binding.Mode = BindingMode.OneWay;
                    binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
                    binding.Converter = new SelectOUTreeBtEnableConverter();
                    BindingOperations.SetBinding(btSelect, Button.IsEnabledProperty, binding);


                    if (!string.IsNullOrWhiteSpace(errorMsg))
                        MessageBox.Show(errorMsg, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    else
                    {
                        DomainOUTreeView.Focus();
                    }
                }));
            }).Start();
        }
        #endregion
    }
}
