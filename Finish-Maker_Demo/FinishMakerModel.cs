using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Finish_Maker_Demo
{
    class FinishMakerModel
    {
        public string UserName { get; set; }
        public int Progress { get; set; }
        public bool ValidateFiles { get; set; }
        public bool ExportLinkCheck { get; set; }
        public bool ProductDataCheck { get; set; }
    }
}
