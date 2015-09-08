using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using COMTools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RCA_Monitor
{
    class Reporter
    {
        COMTools.OutlookTools oTools;
        Outlook.Folder RCAFolder { get; private set; }

        public Reporter()
        {
            oTools = new OutlookTools();
            GetRCAFolder();
        }

        static void Main(string[] args)
        {
            Reporter reporter = new Reporter();
        }

        private void GetRCAFolder()
        {
            Outlook.Store oStore = oTools.oStore;
            Outlook.Folders folders = oStore.GetRootFolder().Folders;
            foreach(Outlook.Folder folder in folders){
                if (folder.Name == "RCAs") { RCAFolder = folder; }
            }
        }
    }
}
