using CreateUsers.Model;
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

namespace CreateUsers.DialogWindows
{
    /// <summary>
    /// Логика взаимодействия для DialogAddEditTemplate.xaml
    /// </summary>
    public partial class DialogAddEditTemplate : Window
    {
        #region Поля
        private DirectoryEntry _ADSession; // Сессия с АД
        private string _pathToPluginDir; // Полный путь к папке плагина
        private ObservableCollection<string> _groups; // Список групп шаблона
        private FormModes.Mode _mode; // Режим работы формы
        private List<string> _templates; // Список шаблонов
        private string _newTemplateToAdd; // Новый добавленый город
        private bool _haveSavedData; // Имеются сохраненные данные
        private int _editIndex; // Позиция в списке редактируемого шаблона
        private string _originalNameEditedTemplate; // Исходные данные редактируемого города
        private List<string> _originalGroupsEditedTemplate; // Список шаблонов
        #endregion

        #region Свойства
        // Строка с новосозданными шаблонами
        public string NewTemplateToAdd
        {
            get { return _newTemplateToAdd; }
            private set { _newTemplateToAdd = value; }
        }
        #endregion

        // Конструктор
        public DialogAddEditTemplate(string pathToPluginDir, List<string> templates, DirectoryEntry entry, FormModes.Mode mode, int editIndex = -1)
        {
            InitializeComponent();
            _pathToPluginDir = pathToPluginDir;
            _mode = mode;
            _templates = templates;
            _ADSession = entry;
            _newTemplateToAdd = "";
            _groups = new ObservableCollection<string>();
            groups.ItemsSource = _groups;
            _haveSavedData = false;
            if (mode == FormModes.Mode.edit)
            {
                Title = "Редактирование шаблона";
                _editIndex = editIndex;
                chCreateMoreTemplates.Visibility = Visibility.Hidden;
                if (editIndex >= 0)
                {
                    _originalNameEditedTemplate = templateName.Text = _templates[editIndex];
                    try
                    {
                        string dataFull = "";
                        using (StreamReader sr = new StreamReader(_pathToPluginDir + "\\templates\\" + _originalNameEditedTemplate + ".txt", Encoding.GetEncoding("windows-1251")))
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
                            _originalGroupsEditedTemplate = _groups.ToList();
                        }
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("Не удалось загрузить шаблон:\r\n" + exp.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
            Regex regex = new Regex(@"[\\\/:\*\?<>\|""]");
            #region Режим создания шаблона
            if (_mode == FormModes.Mode.add)
            {
                try
                {
                    #region Проверка входных данных
                    if (string.IsNullOrWhiteSpace(templateName.Text))
                    {
                        MessageBox.Show("Не указано название шаблона!!!\r\n", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    if (regex.IsMatch(templateName.Text))
                    {
                        MessageBox.Show("Название содержит недопустимые символы!!!\r\n", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    if (_templates.Count > 0)
                    {
                        for (int i = 0; i < _templates.Count; i++)
                        {
                            if (_templates[i].ToUpper() == templateName.Text.Trim().ToUpper())
                            {
                                MessageBox.Show("Шаблон с таким названием уже есть!!\r\nИзмените название или выберите существующий шаблон для редактирования.", "Внимание!!!", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }
                        }
                    }
                    #endregion
                    #region Формирование строки для записи в файл
                    string groupsStr = "";
                    if (groups.Items.Count > 0)
                    {
                        for (int i = 0; i < groups.Items.Count; i++)
                        {
                            groupsStr += groups.Items[i].ToString() + "\r\n";
                        }
                    }
                    #endregion
                    #region Запись новых данных в файл
                    string templateFilePath = _pathToPluginDir + "\\templates\\" + templateName.Text + ".txt";
                    if (!File.Exists(templateFilePath))
                    {
                        File.WriteAllText(templateFilePath, groupsStr, Encoding.GetEncoding("windows-1251"));
                    }
                    else
                    {
                        MessageBox.Show("Шаблон с таким именем уже существует,\r\nукажите другое имя шаблона!!!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    #endregion
                    if (!string.IsNullOrWhiteSpace(NewTemplateToAdd))
                        NewTemplateToAdd += "\r\n" + templateName.Text;
                    else
                        NewTemplateToAdd = templateName.Text;
                    if (chCreateMoreTemplates.IsChecked == true)
                    {
                        templateName.Text = "";
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
                if (_originalNameEditedTemplate != templateName.Text ||
                    !_groups.ToList().SequenceEqual(_originalGroupsEditedTemplate))
                {
                    try
                    {
                        #region Проверка входных данных
                        if (string.IsNullOrWhiteSpace(templateName.Text))
                        {
                            MessageBox.Show("Не указано название шаблона!!!\r\n", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                        if (regex.IsMatch(templateName.Text))
                        {
                            MessageBox.Show("Название содержит недопустимые символы!!!\r\n", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                        #endregion
                        #region Формирование строки для записи в файл
                        string groupsStr = "";
                        if (groups.Items.Count > 0)
                        {
                            for (int i = 0; i < groups.Items.Count; i++)
                            {
                                groupsStr += groups.Items[i].ToString() + "\r\n";
                            }
                        }
                        #endregion
                        #region Запись новых данных в файл
                        string templateFilePath = _pathToPluginDir + "\\templates\\" + templateName.Text + ".txt";
                        if (_originalNameEditedTemplate != templateName.Text)
                        {
                            for (int i = 0; i < _templates.Count; i++)
                            {
                                if (templateName.Text == _templates[i])
                                {
                                    MessageBox.Show("Было изменено имя шаблона.\r\nШаблон с таким имененм уже существует!!!\r\nИзмените имя шаблона. ", "!", MessageBoxButton.OK, MessageBoxImage.Error);
                                    return;
                                }
                            }
                            // Переименование файла шаблона
                            File.Move(_pathToPluginDir + "\\templates\\" + _originalNameEditedTemplate + ".txt", templateFilePath);
                            // Запись новых данных
                            File.WriteAllText(templateFilePath, groupsStr, Encoding.GetEncoding("windows-1251"));
                        }
                        else
                        {
                            File.WriteAllText(templateFilePath, groupsStr, Encoding.GetEncoding("windows-1251"));
                        }
                        #endregion
                        NewTemplateToAdd = templateName.Text;
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
