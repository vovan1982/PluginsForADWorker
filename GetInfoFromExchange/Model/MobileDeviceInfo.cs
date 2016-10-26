using System;

namespace GetInfoFromExchange.Model
{
    public class MobileDeviceInfo
    {
        private string _deviceType;
        private string _deviceModel;
        private string _deviceName;
        private string _deviceUserAgent;
        private DateTime? _lastSyncTime;
        private DateTime? _lastSuccessSyncTime;

        public DateTime? LastSuccessSyncTime
        {
            get { return _lastSuccessSyncTime; }
            set { _lastSuccessSyncTime = value; }
        }

        public DateTime? LastSyncTime
        {
            get { return _lastSyncTime; }
            set { _lastSyncTime = value; }
        }

        public string DeviceUserAgent
        {
            get { return _deviceUserAgent; }
            set { _deviceUserAgent = value; }
        }

        public string DeviceName
        {
            get { return _deviceName; }
            set { _deviceName = value; }
        }

        public string DeviceModel
        {
            get { return _deviceModel; }
            set { _deviceModel = value; }
        }

        public string DeviceType
        {
            get { return _deviceType; }
            set { _deviceType = value; }
        }
    }
}
