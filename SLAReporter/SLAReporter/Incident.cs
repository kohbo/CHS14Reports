using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SLAReporter
{
    public class Incident
    {
        public string number { get; private set; }
        public string facility { get; private set; }
        public string group { get; private set; }
        public string product { get; private set; }
        public string summary { get; private set; }
        public string priority { get; private set; }
        public DateTime submitted { get; private set; }
        public string status { get; private set; }
        public DateTime last_resolved { get; private set; }
        public double days_open { get; private set; }

        public Incident(
                string n,
                string com,
                string g,
                string pro,
                string sum,
                string pri,
                DateTime sub,
                string sta,
                DateTime res,
                double days
            )
        {
            this.number = n;
            this.facility = com;
            this.group = g;
            this.product = pro;
            this.summary = sum;
            this.priority = pri;
            this.submitted = sub;
            this.status = sta;
            this.last_resolved = res;
            this.days_open = days;
        }
    }
}
