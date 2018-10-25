using OutlookAddIn1.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    public class ResourceSaver : IStateSaver
    {
        public bool IsEnabled { get; set; }

        public void Load()
        {
            IsEnabled = Settings.Default.EnabledState;
        }

        public void Save()
        {
            Settings.Default.EnabledState = IsEnabled;
        }
    }
}
