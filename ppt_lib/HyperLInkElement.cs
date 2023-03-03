using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppt_lib
{
    internal class HyperLInkElement
    {
        public string id="";
        public Uri url =null;

        public HyperLInkElement(Uri url ,string id)
        {
            this.url=url;
            this.id =   id ;


        }
        
        
    }
}
