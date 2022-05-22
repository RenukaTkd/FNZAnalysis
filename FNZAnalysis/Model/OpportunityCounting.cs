using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FNZAnalysis.Model
{
    public class OpportunityCounting
    {
        public int red_aua { get; set; }
        public string red_name { get; set; }
        public int red_averagebasispoint { get; set; }
        public EntityReference red_division { get; set; }
        public int red_opportunitycountingid { get; set; }
        public int red_includeinrevenueforecast { get; set; }
        public int red_probability { get; set; }
        public int red_stage { get; set; }
        public int red_targetsigningdate { get; set; }
        public int red_totalscopingconsultancyfeemillion { get; set; }
        public int red_totaldefeemillion { get; set; }
    }
}
