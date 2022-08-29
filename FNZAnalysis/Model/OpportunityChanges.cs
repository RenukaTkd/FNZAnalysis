using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FNZAnalysis.Model
{
   public class OpportunityChanges
    {
        public string SummaryDescription { get; set; }
        public string opportunitycode { get; set; }
        public string ElementChanged { get; set; }

        public string PreviousInput { get; set; }
        public string PresentInput { get; set; }

        public string Division { get; set; }

        public string Modifiedon { get; set; }
        public string CustomerName { get; set; }

        public string Modifiedby { get; set; }
        public string DealOwner { get; set; }

        public string SLTMember { get; set; }

    }
}
