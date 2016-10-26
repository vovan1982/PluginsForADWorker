using GetInfoFromExchange.Model;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace GetInfoFromExchange
{
    /// <summary>
    /// Логика взаимодействия для InboxMessageRule.xaml
    /// </summary>
    public partial class InboxMessageRule : Window
    {

        // Конструктор
        public InboxMessageRule(List<InboxMessageRuleInfo> inboxMessageRule)
        {
            InitializeComponent();
            if (inboxMessageRule.Count > 0 && inboxMessageRule[0] != null)
            {
                inboxMessageRule.Sort(new InboxMessageRuleInfoComparer());
                for (int i = 0; i < inboxMessageRule.Count; i++)
                {
                    Expander newExpander = createExpanderConnect(inboxMessageRule[i].Name, inboxMessageRule[i].Description);
                    if (!inboxMessageRule[i].IsEnabled)
                        newExpander.Foreground = Brushes.Red;
                    rootStackPanel.Children.Add(newExpander);
                }
            }
        }

        // Создание разворачиваемого правила
        private Expander createExpanderConnect(string header, string data)
        {
            Expander dynamicExpander = new Expander();
            dynamicExpander.Header = header;
            dynamicExpander.Foreground = Brushes.Green;
            Border dinamicBorder = new Border();
            dinamicBorder.BorderBrush = Brushes.Black;
            dinamicBorder.BorderThickness = new Thickness(1);
            dinamicBorder.Height = 350;
            dinamicBorder.Margin = new Thickness(10, 0, 8, 0);
            ScrollViewer dinamicScrollViewer = new ScrollViewer();
            dinamicScrollViewer.Margin = new Thickness(1);
            TextBox dinamicTextBox = new TextBox();
            dinamicTextBox.IsReadOnly = true;
            dinamicTextBox.TextWrapping = TextWrapping.Wrap;
            dinamicTextBox.Foreground = Brushes.Black;
            dinamicTextBox.Text = data;
            dinamicScrollViewer.Content = dinamicTextBox;
            dinamicBorder.Child = dinamicScrollViewer;
            dynamicExpander.Content = dinamicBorder;
            return dynamicExpander;
        }
        // Нажата кнопка закрытия окна
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        // Нажата кнопка развернуть все
        private void btExpandAll_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < rootStackPanel.Children.Count; i++)
            {
                Expander child = rootStackPanel.Children[i] as Expander;
                child.IsExpanded = true;
            }
        }
        // Нажата кнопка свернуть все
        private void btCollapseAll_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < rootStackPanel.Children.Count; i++)
            {
                Expander child = rootStackPanel.Children[i] as Expander;
                child.IsExpanded = false;
            }
        }
    }
}
