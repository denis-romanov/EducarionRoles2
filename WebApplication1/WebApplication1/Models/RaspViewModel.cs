using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
//using System.Collections.Generic;
using System.Web.Mvc;

namespace WebApplication1.Models
{
    public class RaspViewModel
    {
        public SelectList YearList { get; set; }
        public string[] SelectedYear { get; set; }

        public SelectList SemsList { get; set; }
        public string[] SelectedSem { get; set; }
    }
    public class Year
    {
        public string y { get; set; }
    }
    public class Sem
    {
        public string s { get; set; }
    }
}