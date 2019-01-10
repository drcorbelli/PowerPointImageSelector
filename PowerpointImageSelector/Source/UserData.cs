using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace PowerpointImageSelector
{
    public class UserData
    {
        public string Title { get; set; }
        public string TextField { get; set; }
        public List<string> Keywords { get; set; }
        public List<Image> Images { get; set; }

        public UserData()
        {
            Keywords = new List<string>();
            Images = new List<Image>();
        }
    }

}
