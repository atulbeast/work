using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides.Model
{
    class IssuesModel
    {

        public string Issues { get; set; }
        public string Owner { get; set; }
        public string Date { get; set; } = DateTime.Today.ToShortDateString();
    }
}
