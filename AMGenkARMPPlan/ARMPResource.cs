using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    class ARMPResource
    {
        public string name { get; set; }
        public string amei { get; set; }

        public ARMPResource(string name, string amei)
        {
            this.name = name;
            this.amei = amei;
        }

    }
}
