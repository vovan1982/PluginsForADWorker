using System;
using System.Windows;
using System.Windows.Controls;

namespace DismissEmployee.UISearchTextBox
{
    class ListBoxEx : ListBox
    {
        public enum ItemStyles
        {
            NormalStyle,
            CheckBoxStyle,
            RadioBoxStyle
        }
        private ItemStyles m_extendedStyle;

        public ItemStyles ExtendedStyle
        {
            get { return m_extendedStyle; }
            set
            {
                m_extendedStyle = value;

                // load resources
                ResourceDictionary resDict = new ResourceDictionary();
                resDict.Source = new Uri("pack://application:,,,/DismissEmployee;component/UISearchTextBox/Themes/ListBoxEx.xaml");
                if (resDict.Source == null)
                    throw new SystemException();

                switch (value)
                {
                    case ItemStyles.NormalStyle:
                        this.Style = (Style)resDict["NormalListBox"];
                        break;
                    case ItemStyles.CheckBoxStyle:
                        this.Style = (Style)resDict["CheckListBox"];
                        break;
                    case ItemStyles.RadioBoxStyle:
                        this.Style = (Style)resDict["RadioListBox"];
                        break;
                }
            }
        }

        static ListBoxEx()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(ListBoxEx), new FrameworkPropertyMetadata(typeof(ListBoxEx)));
        }

        public ListBoxEx(ItemStyles style)
            : base()
        {
            ExtendedStyle = style;
        }

        public ListBoxEx()
            : base()
        {
            ExtendedStyle = ItemStyles.NormalStyle;
        }
    }
}
