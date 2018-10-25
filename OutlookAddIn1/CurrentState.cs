using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OutlookAddIn1
{
    public class CurrentState : IStateSaver
    {
        //private string location = Directory.GetCurrentDirectory() + @"/data.xml";

        public bool IsEnabled { get; set; }
        
        XmlDocument doc = new XmlDocument();
        XmlNode EnabledNode;
        public CurrentState()
        {
            FileInfo file = new FileInfo(location);
            if(!file.Exists)
            {
                file.Create();
                doc.LoadXml("<root><enabled state=\"false\"/></root>");
                doc.Save(location);
            }
           
                doc.Load(location);
          
            EnabledNode = doc.SelectSingleNode("root/enabled");

        }
        public void Save()
        {
            EnabledNode.Attributes["state"].Value = IsEnabled.ToString();
            doc.Save(location);
            Console.WriteLine("saved");
        }

        public void Load()
        {
            IsEnabled = bool.Parse(EnabledNode.Attributes["state"].Value);
        }
    }
}
