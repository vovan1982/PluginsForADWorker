using CreateUsersPPK.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace CreateUsersPPK.DialogWindows
{
    /// <summary>
    /// Логика взаимодействия для DialogAddEditCity.xaml
    /// </summary>
    public partial class DialogAddEditCity : Window
    {
        #region Поля
        private DirectoryEntry _ADSession; // Сессия с АД
        private string _pathToPluginDir; // Полный путь к папке плагина
        private ObservableCollection<string> _groups; // Список групп в которых должен состоять пользователь
        private FormModes.Mode _mode; // Режим работы формы
        private List<Cities> _сities; // Список городов
        private string _newCityToAdd; // Новый добавленый город
        private bool _haveSavedData; // Имеются сохраненные данные
        private int _editIndex; // Позиция в списке редактируемого города
        private Cities _originalDataEditedCity; // Исходные данные редактируемого города
        #endregion

        #region Свойства
        // Строка с новосозданными городами
        public string NewCityToAdd
        {
            get { return _newCityToAdd; }
            private set { _newCityToAdd = value; }
        }
        #endregion

        // Конструктор
        public DialogAddEditCity(string pathToPluginDir, List<Cities> сities, DirectoryEntry entry,FormModes.Mode mode, int editIndex = -1)
        {
            InitializeComponent();
            _pathToPluginDir = pathToPluginDir;
            _mode = mode;
            _сities = сities;
            _ADSession = entry;
            _newCityToAdd = "";
            _groups = new ObservableCollection<string>();
            groups.ItemsSource = _groups;
            _haveSavedData = false;

            if (mode == FormModes.Mode.edit)
            {
                Title = "Редактирование города";
                _editIndex = editIndex;
                _originalDataEditedCity = _сities[editIndex];
                chCreateSeveralCities.Visibility = Visibility.Hidden;
                if (editIndex >= 0)
                {
                    displayName.Text = _сities[editIndex].DisplayName;
                    name.Text = _сities[editIndex].Name;
                    adress.Text = _сities[editIndex].Adress;
                    if (_сities[editIndex].Groups.Count > 0)
                    {
                        for (int i = 0; i < _сities[editIndex].Groups.Count; i++)
                        {
                            _groups.Add(_сities[editIndex].Groups[i]);
                        }
                    }
                }
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
                        if (!isExist)
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
                foreach (var column in groupsGridView.Columns)
                {
                    column.Width = column.ActualWidth;
                    column.Width = Double.NaN;
                }
            }
        }
        // Нажата кнопка удаления групп
        private void btDelGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                while (groups.SelectedItems.Count > 0)
                {
                    _groups.Remove((string)groups.SelectedItems[0]);
                }
                if (groups.Items.Count > 0)
                {
                    groups.SelectedIndex = 0;
                    groups.Focus();
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Нажата кнопка закрытия окна
        private void btClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = _haveSavedData;
            Close();
        }
        // Нажата кнопка сохранения
        private void btSave_Click(object sender, RoutedEventArgs e)
        {
            #region Режим добавления нового города
            if (_mode == FormModes.Mode.add)
            {
                try
                {
                    #region Проверка входных данных
                    if (string.IsNullOrWhiteSpace(displayName.Text) || string.IsNullOrWhiteSpace(name.Text))
                    {
                        MessageBox.Show("Не все поля заполнены!!!\r\n", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    if (_сities.Count > 0)
                    {
                        for (int i = 0; i < _сities.Count; i++)
                        {
                            if (_сities[i].DisplayName.Trim().ToUpper() == displayName.Text.Trim().ToUpper())
                            {
                                MessageBox.Show("Город с таким названием уже есть!!\r\nИзмените название или выберите существующий город для редактирования.", "Внимание!!!", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }
                        }
                    }
                    #endregion
                    #region Формирование строки для записи в файл
                    string cityString = displayName.Text + ";" + name.Text + ";" + adress.Text + ";";
                    if (groups.Items.Count > 0)
                    {
                        string groupsStr = "";
                        for (int i = 0; i < groups.Items.Count; i++)
                        {
                            groupsStr += groups.Items[i].ToString() + ",";
                        }
                        cityString += groupsStr.TrimEnd(',');
                    }
                    #endregion
                    #region Запись новых данных в файл
                    if (File.Exists(_pathToPluginDir + "\\cities.dat"))
                    {
                        StreamReader sr = new StreamReader(_pathToPluginDir + "\\cities.dat", Encoding.GetEncoding("windows-1251"));
                        string fullData = sr.ReadToEnd();
                        sr.Close();
                        if (fullData.EndsWith("\r\n"))
                        {
                            File.AppendAllText(_pathToPluginDir + "\\cities.dat", cityString + "\r\n", Encoding.GetEncoding("windows-1251"));
                        }
                        else if(string.IsNullOrWhiteSpace(fullData))
                        {
                            File.AppendAllText(_pathToPluginDir + "\\cities.dat", cityString + "\r\n", Encoding.GetEncoding("windows-1251"));
                        }
                        else
                        {
                            File.AppendAllText(_pathToPluginDir + "\\cities.dat", "\r\n" + cityString + "\r\n", Encoding.GetEncoding("windows-1251"));
                        }
                    }
                    else
                    {
                        File.WriteAllText(_pathToPluginDir + "\\cities.dat", cityString + "\r\n", Encoding.GetEncoding("windows-1251"));
                    }
                    #endregion
                    if (!string.IsNullOrWhiteSpace(NewCityToAdd))
                        NewCityToAdd += "\r\n" + cityString;
                    else
                        NewCityToAdd = cityString;
                    if (chCreateSeveralCities.IsChecked == true)
                    {
                        displayName.Text = "";
                        name.Text = "";
                        adress.Text = "";
                        _groups.Clear();
                        _haveSavedData = true;
                    }
                    else
                    {
                        DialogResult = true;
                        Close();
                    }
                }
                catch (Exception exp)
                {
                    MessageBox.Show("В процессе сохранения возникла ошибка:\r\n" + exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            #endregion

            #region Режим редактирования города
            if (_mode == FormModes.Mode.edit)
            {
                if (_originalDataEditedCity.DisplayName != displayName.Text ||
                    _originalDataEditedCity.Name != name.Text ||
                    _originalDataEditedCity.Adress != adress.Text ||
                    !_groups.ToList().SequenceEqual(_originalDataEditedCity.Groups))
                {
                    try
                    {
                        #region Формирование строки со старыми данными
                        string oldData = _originalDataEditedCity.DisplayName + ";" + _originalDataEditedCity.Name + ";" + _originalDataEditedCity.Adress + ";";
                        if (_originalDataEditedCity.Groups.Count > 0)
                        {
                            string groupsStr = "";
                            for (int i = 0; i < groups.Items.Count; i++)
                            {
                                groupsStr += _originalDataEditedCity.Groups[i].ToString() + ",";
                            }
                            oldData += groupsStr.TrimEnd(',');
                        }
                        #endregion
                        #region Формирование строки с новыми данными
                        string newData = displayName.Text + ";" + name.Text + ";" + adress.Text + ";";
                        if (groups.Items.Count > 0)
                        {
                            string groupsStr = "";
                            for (int i = 0; i < groups.Items.Count; i++)
                            {
                                groupsStr += groups.Items[i].ToString() + ",";
                            }
                            newData += groupsStr.TrimEnd(',');
                        }
                        #endregion
                        #region Запись измененных данных
                        if (_originalDataEditedCity.DisplayName != displayName.Text)
                        {
                            for (int i = 0; i < _сities.Count; i++)
                            {
                                if (displayName.Text == _сities[i].DisplayName)
                                {
                                    MessageBox.Show("Было изменено отображаемое имя.\r\nГород с таким имененм уже существует в списке!!!\r\nИзмените отображаемое имя. ", "!", MessageBoxButton.OK, MessageBoxImage.Error);
                                    return;
                                }
                            }
                            StreamReader reader = new StreamReader(_pathToPluginDir + "\\cities.dat", Encoding.GetEncoding("windows-1251"));
                            string content = reader.ReadToEnd();
                            reader.Close();

                            content = Regex.Replace(content, oldData, newData);

                            StreamWriter writer = new StreamWriter(_pathToPluginDir + "\\cities.dat", false, Encoding.GetEncoding("windows-1251"));
                            writer.Write(content);
                            writer.Close();
                        }
                        else
                        {
                            StreamReader reader = new StreamReader(_pathToPluginDir + "\\cities.dat", Encoding.GetEncoding("windows-1251"));
                            string content = reader.ReadToEnd();
                            reader.Close();

                            content = Regex.Replace(content, oldData, newData);

                            StreamWriter writer = new StreamWriter(_pathToPluginDir + "\\cities.dat", false, Encoding.GetEncoding("windows-1251"));
                            writer.Write(content);
                            writer.Close();
                        }
                        #endregion
                        NewCityToAdd = newData;
                        DialogResult = true;
                        Close();
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("В процессе сохранения возникла ошибка:\r\n" + exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            #endregion
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
