using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Reflection;

namespace COMTools
{
    public class OutlookTools
    {
        public Outlook.Application oApp
        {
            get
            {
                return GetApplicationObject();
            }
        }

        public Outlook.Store oStore
        {
            private set
            {
                oStore = value;
            }
            get
            {
                if (oStore == null)
                {
                    return oApp.Session.GetStoreFromID(GetStoreID(oApp.Session.Stores));
                }
                return oStore;
            }
        }

        private Outlook.Application GetApplicationObject()
        {

            Outlook.Application application = null;

            // Check if there is an Outlook process running. 
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {

                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object. 
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            else
            {

                // If not, create a new instance of Outlook and log on to the default profile. 
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "", Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object. 
            return application;
        }

        private string GetStoreID(Outlook.Stores stores)
        {
            Console.WriteLine("Enter the number corresponding with the data file...");
            int index = 0;
            foreach (Outlook.Store store in stores)
            {
                Console.WriteLine(index + ": " + store.DisplayName);
                index++;
            }
            Console.Write("Selection > ");
            int sel = Convert.ToInt32(Console.Read()) - 47;
            return stores[sel].StoreID;
        }
    }
}
