using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace archlabab9
{
    internal class TitlePageData
    {
        public string Discilpline { get; set; }
        public string Title { get; set; }

        public string Teacher { get; set; }
        public string WorkType { get; set; }
        public string WorkNumber { get; set; }
        public string Year { get; set; } = DateTime.Now.Year.ToString();

    }
}
