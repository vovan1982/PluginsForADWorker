﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace AutoReplyAndForwarding.UISearchTextBox
{
    public class SearchEventArgs : RoutedEventArgs
    {
        private string m_keyword = "";

        public string Keyword
        {
            get { return m_keyword; }
            set { m_keyword = value; }
        }
        private List<string> m_sections = new List<string>();

        public List<string> Sections
        {
            get { return m_sections; }
            set { m_sections = value; }
        }
        public SearchEventArgs()
            : base()
        {

        }
        public SearchEventArgs(RoutedEvent routedEvent)
            : base(routedEvent)
        {

        }
    }

    public class SearchTextBox : TextBox
    {
        public static DependencyProperty LabelTextProperty =
            DependencyProperty.Register(
                "LabelText",
                typeof(string),
                typeof(SearchTextBox));

        public static DependencyProperty LabelTextColorProperty =
            DependencyProperty.Register(
                "LabelTextColor",
                typeof(Brush),
                typeof(SearchTextBox));

        private static DependencyPropertyKey HasTextPropertyKey =
            DependencyProperty.RegisterReadOnly(
                "HasText",
                typeof(bool),
                typeof(SearchTextBox),
                new PropertyMetadata());
        public static DependencyProperty HasTextProperty = HasTextPropertyKey.DependencyProperty;

        private static DependencyPropertyKey IsMouseLeftButtonDownPropertyKey =
            DependencyProperty.RegisterReadOnly(
                "IsMouseLeftButtonDown",
                typeof(bool),
                typeof(SearchTextBox),
                new PropertyMetadata());
        public static DependencyProperty IsMouseLeftButtonDownProperty = IsMouseLeftButtonDownPropertyKey.DependencyProperty;

        public static readonly RoutedEvent SearchEvent =
            EventManager.RegisterRoutedEvent(
                "Search",
                RoutingStrategy.Bubble,
                typeof(RoutedEventHandler),
                typeof(SearchTextBox));

        static SearchTextBox()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(SearchTextBox),
                new FrameworkPropertyMetadata(typeof(SearchTextBox)));
        }

        public SearchTextBox()
            : base()
        {

        }

        protected override void OnTextChanged(TextChangedEventArgs e)
        {
            base.OnTextChanged(e);

            HasText = Text.Length != 0;
            HidePopup();
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn protected override void OnMouseDown(MouseButtonEventArgs e)
        ///
        /// @brief  Override the default method. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ///
        /// @param  e   Event information to send to registered event handlers. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);

            // if users click on the editing area, the pop up will be hidden
            Type sourceType = e.OriginalSource.GetType();
            if (sourceType != typeof(Image))
                HidePopup();
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            this.MouseLeave += new MouseEventHandler(SearchTextBox_MouseLeave);
            Border iconBorder = GetTemplateChild("PART_SearchIconBorder") as Border;
            if (iconBorder != null)
            {
                iconBorder.MouseLeftButtonDown += new MouseButtonEventHandler(IconBorder_MouseLeftButtonDown);
                iconBorder.MouseLeftButtonUp += new MouseButtonEventHandler(IconBorder_MouseLeftButtonUp);
                iconBorder.MouseLeave += new MouseEventHandler(IconBorder_MouseLeave);
                iconBorder.MouseDown += new MouseButtonEventHandler(SearchIcon_MouseDown);
            }

            int size = 0;
            if (ShowSectionButton)
            {
                iconBorder = GetTemplateChild("PART_SpecifySearchType") as Border;
                if (iconBorder != null)
                {
                    iconBorder.MouseDown += new MouseButtonEventHandler(ChooseSection_MouseDown);
                }
                size = 15;
            }
            Image iconChoose = GetTemplateChild("SpecifySearchType") as Image;
            if (iconChoose != null)
                iconChoose.Width = iconChoose.Height = size;

            iconBorder = GetTemplateChild("PART_PreviousItem") as Border;
            if (iconBorder != null)
                iconBorder.MouseDown += new MouseButtonEventHandler(PreviousItem_MouseDown);

            //////////////////////////////////////////////////////////////////////////
        }

        public void SetFocus()
        {
            this.Focus();
        }

        void SearchIcon_MouseDown(object sender, MouseButtonEventArgs e)
        {
            HidePopup();
            if (this.HasText)
            {
                this.Clear();
            }
        }



        void SearchTextBox_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!m_listPopup.IsMouseOver)
                HidePopup();
        }




        private void IconBorder_MouseLeftButtonDown(object obj, MouseButtonEventArgs e)
        {
            IsMouseLeftButtonDown = true;

        }

        private void IconBorder_MouseLeftButtonUp(object obj, MouseButtonEventArgs e)
        {
            if (!IsMouseLeftButtonDown) return;

            if (HasText)
            {
                RaiseSearchEvent();
            }

            IsMouseLeftButtonDown = false;
        }

        private void IconBorder_MouseLeave(object obj, MouseEventArgs e)
        {
            IsMouseLeftButtonDown = false;

        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Text = "";
            }
            else if ((e.Key == Key.Return || e.Key == Key.Enter))
            {
                RaiseSearchEvent();
            }
            else
            {
                base.OnKeyDown(e);
            }
        }

        private void RaiseSearchEvent()
        {
            if (this.Text == "")
                return;
            if (!m_listPreviousItem.Items.Contains(this.Text))
            {
                m_listPreviousItem.Items.Insert(0, this.Text);
                //m_listPreviousItem.Items.Add(this.Text);
            }
            else
            {
                int index = m_listPreviousItem.Items.IndexOf(this.Text);
                m_listPreviousItem.Items.RemoveAt(index);
                m_listPreviousItem.Items.Insert(0, this.Text);
            }

            SearchEventArgs args = new SearchEventArgs(SearchEvent);
            args.Keyword = this.Text;
            if (m_listSection != null)
            {
                args.Sections = (List<string>)m_listSection.SelectedItems.Cast<string>().ToList();
            }
            RaiseEvent(args);
        }

        public string LabelText
        {
            get { return (string)GetValue(LabelTextProperty); }
            set { SetValue(LabelTextProperty, value); }
        }

        public Brush LabelTextColor
        {
            get { return (Brush)GetValue(LabelTextColorProperty); }
            set { SetValue(LabelTextColorProperty, value); }
        }

        public bool HasText
        {
            get { return (bool)GetValue(HasTextProperty); }
            private set { SetValue(HasTextPropertyKey, value); }
        }

        public bool IsMouseLeftButtonDown
        {
            get { return (bool)GetValue(IsMouseLeftButtonDownProperty); }
            private set { SetValue(IsMouseLeftButtonDownPropertyKey, value); }
        }

        public event RoutedEventHandler OnSearch
        {
            add { AddHandler(SearchEvent, value); }
            remove { RemoveHandler(SearchEvent, value); }
        }

        #region Stuff added by Le Duc Anh

        public static DependencyProperty SectionsListProperty =
            DependencyProperty.Register(
                "SectionsList",
                typeof(List<string>),
                typeof(SearchTextBox),
                new FrameworkPropertyMetadata(null,
                                                FrameworkPropertyMetadataOptions.None)
             );

        public List<string> SectionsList
        {
            get { return (List<string>)GetValue(SectionsListProperty); }
            set
            {
                SetValue(SectionsListProperty, value);

            }
        }

        public List<string> SelectedSections
        {
            get { return (List<string>)m_listSection.SelectedItems.Cast<string>().ToList(); }
        }

        private bool m_showSectionButton = true;

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @property   public bool ShowSectionButton
        ///
        /// @brief  Gets or sets a value indicating whether the choose section button is shown. 
        ///
        /// @return true if show choose section button, false if not. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        public bool ShowSectionButton
        {
            get { return m_showSectionButton; }
            set
            {
                m_showSectionButton = value;
            }
        }

        public enum SectionsStyles
        {
            NormalStyle,
            CheckBoxStyle,
            RadioBoxStyle
        }
        private SectionsStyles m_itemStyle = SectionsStyles.CheckBoxStyle;

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @property   public SectionsStyles ItemStyle
        ///
        /// @brief  Gets or sets the style's of each displayed section. This property is valid when
        ///         ShowSectionButton is set to true. 
        ///
        /// @return The item style. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        public SectionsStyles SectionsStyle
        {
            get { return m_itemStyle; }
            set { m_itemStyle = value; }
        }
        private Popup m_listPopup = new Popup();
        private ListBoxEx m_listSection = null;
        private ListBoxEx m_listPreviousItem = null;

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn private void BuildPopup()
        ///
        /// @brief  Builds the pop up and related items
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        private void BuildPopup()
        {
            // initialize the pop up
            m_listPopup.PopupAnimation = PopupAnimation.Fade;
            m_listPopup.Placement = PlacementMode.Relative;
            m_listPopup.PlacementTarget = this;
            m_listPopup.PlacementRectangle = new Rect(0, this.ActualHeight, 30, 30);
            m_listPopup.Width = this.ActualWidth;
            // initialize the sections' list
            if (ShowSectionButton)
            {
                m_listSection = new ListBoxEx((int)m_itemStyle + ListBoxEx.ItemStyles.NormalStyle);

                //////////////////////////////////////////////////////////////////////////
                m_listSection.Items.Clear();
                // add items into the list
                // is there any smarter way?
                if (SectionsList != null)
                {
                    foreach (string item in SectionsList)
                        m_listSection.Items.Add(item);
                    if (m_itemStyle == SectionsStyles.RadioBoxStyle)
                        m_listSection.SelectedIndex = 0;
                }
                //////////////////////////////////////////////////////////////////////////

                m_listSection.Width = this.Width;
                m_listSection.MouseLeave += new MouseEventHandler(ListSection_MouseLeave);

            }

            // initialize the previous items' list
            m_listPreviousItem = new ListBoxEx();
            m_listPreviousItem.MouseLeave += new MouseEventHandler(ListPreviousItem_MouseLeave);
            m_listPreviousItem.SelectionChanged += new SelectionChangedEventHandler(ListPreviousItem_SelectionChanged);
            m_listPreviousItem.Width = this.Width;
            m_listPreviousItem.Height = 85;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn private void HidePopup()
        ///
        /// @brief  Hides the pop up. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        private void HidePopup()
        {
            m_listPopup.IsOpen = false;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn private void ShowPopup(UIElement item)
        ///
        /// @brief  Shows the pop up. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ///
        /// @param  item    The item. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        private void ShowPopup(UIElement item)
        {
            m_listPopup.StaysOpen = true;

            m_listPopup.Child = item;
            m_listPopup.IsOpen = true;

        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn void ListPreviousItem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        ///
        /// @brief  Event handler. Called by m_listPreviousItem for selection changed events. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ///
        /// @param  sender  Source of the event. 
        /// @param  e       Selection changed event information. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        void ListPreviousItem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // fetch the previous keyword into the text box
            if (m_listPreviousItem.SelectedItems.Count > 0)
            {
                this.Text = m_listPreviousItem.SelectedItems[0].ToString();
                // close the pop up
                HidePopup();

                this.SetFocus();
                this.SelectionStart = this.Text.Length;
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn void ListPreviousItem_MouseLeave(object sender, MouseEventArgs e)
        ///
        /// @brief  Event handler. Called by m_listPreviousItem for mouse leave events. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ///
        /// @param  sender  Source of the event. 
        /// @param  e       Mouse event information. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        void ListPreviousItem_MouseLeave(object sender, MouseEventArgs e)
        {
            // close the pop up
            HidePopup();
            m_listPreviousItem.SelectedIndex = -1;
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn void PreviousItem_MouseDown(object sender, MouseButtonEventArgs e)
        ///
        /// @brief  Event handler. Called by PreviousItem for mouse down events, showing previously typed
        ///         keywords. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ///
        /// @param  sender  Source of the event. 
        /// @param  e       Mouse button event information. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        void PreviousItem_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (m_listPreviousItem.Items.Count != 0)
                ShowPopup(m_listPreviousItem);
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn void ListSection_MouseLeave(object sender, MouseEventArgs e)
        ///
        /// @brief  Event handler. Called by m_listSection for mouse leave events. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ///
        /// @param  sender  Source of the event. 
        /// @param  e       Mouse event information. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        void ListSection_MouseLeave(object sender, MouseEventArgs e)
        {
            // close the pop up
            HidePopup();
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////
        /// @fn void ChooseSection_MouseDown(object sender, MouseButtonEventArgs e)
        ///
        /// @brief  Event handler. Called by ChooseSection for mouse down events. 
        ///
        /// @author Le Duc Anh
        /// @date   8/14/2010
        ///
        /// @param  sender  Source of the event. 
        /// @param  e       Mouse button event information. 
        ////////////////////////////////////////////////////////////////////////////////////////////////////

        void ChooseSection_MouseDown(object sender, MouseButtonEventArgs e)
        {

            if (SectionsList == null)
                return;
            if (SectionsList.Count != 0)
                ShowPopup(m_listSection);
        }

        protected override void OnLostFocus(RoutedEventArgs e)
        {
            base.OnLostFocus(e);
            if (!HasText)
                this.LabelText = "Поиск";

            m_listPopup.StaysOpen = false;
        }

        protected override void OnGotFocus(RoutedEventArgs e)
        {
            base.OnGotFocus(e);
            if (!HasText)
                this.LabelText = "";
            m_listPopup.StaysOpen = true;

        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);
            BuildPopup();
        }
        #endregion
    }
}
