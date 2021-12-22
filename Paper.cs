using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ifan.PDF
{
    /// <summary>
    /// 用于存储图纸尺寸设置的类
    /// </summary>
    [Serializable]
    public class Paper
    {
        public double width { get; set; }
        public double height { get; set; }
        public PaperSize Size { get; set; }

        public Paper(PaperSize size,double w, double h)
        {
            width = w;
            height = h;
            Size = size;
        }

        public Paper()
        {
        }
    }


}
