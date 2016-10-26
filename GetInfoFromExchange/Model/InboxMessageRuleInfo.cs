using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GetInfoFromExchange.Model
{
    public class InboxMessageRuleInfo
    {
        private string _name;
        private string _description;
        private bool _isEnabled;
        private int _priority;

        public int Priority
        {
            get { return _priority; }
            set { _priority = value; }
        }

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set { _isEnabled = value; }
        }

        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
    }

    public class InboxMessageRuleInfoComparer : IComparer<InboxMessageRuleInfo>
    {
        public int Compare(InboxMessageRuleInfo x, InboxMessageRuleInfo y)
        {
            if (x == null)
            {
                if (y == null)
                {
                    // If x is null and y is null, they're
                    // equal. 
                    return 0;
                }
                else
                {
                    // If x is null and y is not null, y
                    // is greater. 
                    return -1;
                }
            }
            else
            {
                // If x is not null...
                //
                if (y == null)
                // ...and y is null, x is greater.
                {
                    return 1;
                }
                else
                {
                    // ...and y is not null, compare the 
                    // lengths of the two strings.
                    //
                    if (x.Priority > y.Priority)
                        return 1;
                    else if (x.Priority < y.Priority)
                        return -1;
                    else
                        return 0;
                }
            }

        }
    }
}
