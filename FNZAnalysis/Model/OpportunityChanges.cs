using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FNZAnalysis.Model
{
   public class OpportunityChanges
    {
        public string OpportunityName { get; set; }
        public string ElementChanged { get; set; }

        public string PreviousInput { get; set; }
        public string PresentInput { get; set; }

        public string Division { get; set; }

    }
}
