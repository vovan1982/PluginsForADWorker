using GetInfoFromExchange.Model;
using System.Collections.Generic;
using System.Windows;

namespace GetInfoFromExchange
{
    /// <summary>
    /// Логика взаимодействия для MobileDevice.xaml
    /// </summary>
    public partial class MobileDevice : Window
    {
        public MobileDevice(List<MobileDeviceInfo> mobileDeviceData)
        {
            InitializeComponent();
            if (mobileDeviceData.Count > 0 && mobileDeviceData[0] != null)
            {
                rootDataGrid.ItemsSource = mobileDeviceData;
                if (rootDataGrid.Items.Count > 0)
                    rootDataGrid.SelectedIndex = 0;
                rootDataGrid.Focus();
            }
        }
    }
}
