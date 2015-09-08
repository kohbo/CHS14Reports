using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using BMC.ARSystem;

namespace Queue_Monitor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            BMC.ARSystem.Server arserver = new BMC.ARSystem.Server();
            arserver.Login("10.244.0.40", "juan.menendez", "Pa55w0rd", "");
            String RequestID = "000002973185";

            string FromForm = ((BMC.ARSystem.EntryDescription)arserver.GetListEntry("HPD:Help Desk", string.Format("'1' = \"{0}\"", RequestID))[0]).Description;

            string qualification = string.Format("'1' = " + RequestID);

            BMC.ARSystem.EntryListFieldList fieldList = new BMC.ARSystem.EntryListFieldList();
            fieldList.Add(new BMC.ARSystem.EntryListField(8));
            fieldList.Add(new BMC.ARSystem.EntryListField(3));

            var entryList = arserver.GetListEntryWithFields("HPD:Help Desk", qualification, fieldList, 0, 0);

            Console.WriteLine(entryList[0].FieldValues[8]);
            Console.WriteLine(entryList[0].FieldValues[3]);

            Console.ReadLine();
        }
    }
}
