using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    public class TextSaver : IStateSaver
    {
        private bool _IsEnabled;
        private string location = Directory.GetCurrentDirectory() + @"/data.txt";

        public bool IsEnabled { get { return _IsEnabled; } set { _IsEnabled = value; } }
        public TextSaver()
        {
            if(!File.Exists(location))
            {
                Save();
            }
            Load();            
        }
        public void Load()
        {
            var result = bool.Parse(File.ReadAllText(location));

        }

        public void Save()
        {

            Debug.WriteLine($"saved {IsEnabled}");
            File.WriteAllText(location, IsEnabled.ToString());
        }
    }
}
