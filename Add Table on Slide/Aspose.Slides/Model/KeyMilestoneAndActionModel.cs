using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides.Model
{
    class KeyMilestoneAndActionModel
    {

        public string Date { get; set; } = DateTime.Today.ToShortDateString();
        public string TaskEvent { get; set; }
        public string Resp { get; set; }
        public string Status { get; set; }

    }
}
