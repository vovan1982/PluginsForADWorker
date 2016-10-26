using System;

namespace GetInfoFromExchange.Model
{
    public class UserAutoReplayInfo 
    {
        private string _state;
        private DateTime _startDateTime;
        private DateTime _endDateTime;
        private string _internalMessageText;
        private string _externalMessageText;

        public string ExternalMessageText
        {
            get { return _externalMessageText; }
            set { _externalMessageText = value; }
        }

        public string InternalMessageText
        {
            get { return _internalMessageText; }
            set { _internalMessageText = value; }
        }

        public DateTime EndDateTime
        {
            get { return _endDateTime; }
            set { _endDateTime = value; }
        }

        public DateTime StartDateTime
        {
            get { return _startDateTime; }
            set { _startDateTime = value; }
        }

        public string State
        {
            get { return _state; }
            set { _state = value; }
        }
    }
}
