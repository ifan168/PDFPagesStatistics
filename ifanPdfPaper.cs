using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ifan.PDF
{

    public class ifanPdfPaper
    {
        public static List<Paper> m_papers = new List<Paper> {
           new Paper(PaperSize.A0,1189,841),
            new Paper(PaperSize.A1,841,594),
             new Paper(PaperSize.A2,594,420),
              new Paper(PaperSize.A3,420,297),
               new Paper(PaperSize.A4,297,210),        };
        public static int intTolerance=5;

        public ifanPdfPaper(PdfPage pdfpage)
        {
            _pdfpage = pdfpage;
            JudgePaperSize(_pdfpage);
        }
        public PdfPage GetPdfPage()
        {
            return _pdfpage;
        }
        /// <summary>
        /// 返回图纸尺寸名
        /// </summary>
        public string SizeName
        {
            get
            {
                if (_paperSizeName.ToString().Contains("plus"))
                {
                    return _paperSizeName.ToString().Replace("plus", "+") + plusLength;
                }
                else
                { return _paperSizeName.ToString().Replace("plus", "+"); }
            }
        }
        /// <summary>
        /// 返回图纸绝对文件路径
        /// </summary>
        public string PaperFilePath
        {
            get
            {
                return PaperFilePath;
            }
        }
        /// <summary>
        /// 返回图纸长度
        /// </summary>
        public double PaperSizeWidth
        {
            get { return _paperSizeWidth; }
        }
        /// <summary>
        /// 返回图纸高度
        /// </summary>
        public double PaperSizeHeight
        {
            get { return _paperSizeHeight; }
        }

        double _paperSizeWidth, _paperSizeHeight;
        PaperSize _paperSizeName;
        private double plusLength;
        PdfPage _pdfpage;
  

        /// <summary>
        /// 得到当前页面的图纸尺寸
        /// </summary>
        /// <param name="pdfpage"></param>
        private void JudgePaperSize(PdfPage pdfpage)
        {
            double W = Math.Max(pdfpage.Width.Millimeter, pdfpage.Height.Millimeter);
            double H = Math.Min(pdfpage.Width.Millimeter, pdfpage.Height.Millimeter);
            _paperSizeWidth = W;
            _paperSizeHeight = H;
            _paperSizeName = SizeToTypeName(W, H);
        }

        /// <summary>
        /// 用于返回图纸尺寸对应的图副
        /// </summary>
        /// <param name="w">图纸长度</param>
        /// <param name="h">图纸高度</param>
        /// <returns></returns>
        private PaperSize SizeToTypeName(double w, double h)
        {
            PaperSize ps = PaperSize.Undefined;
            try
            {
                //容差值为图纸高度的±5%
                Paper paper = m_papers.First(p => Math.Abs(p.height - h) < (p.height * intTolerance/100));
                ps = paper.Size;
                //长度长于标准长度的1.1倍，则认为是加长图框
                if (w > (paper.width * 1.1))
                {
                    plusLength = (_paperSizeWidth / paper.width) - 1;
                    if (paper.Size == PaperSize.A0)
                    {
                        plusLength = Math.Round(plusLength / 0.125) * 0.125;
                        if (Math.Abs((plusLength + 1) * paper.width - w) > 0.125 * 0.1 * paper.width)
                        {
                            plusLength += 0.125;
                        }
                    }
                    else
                    {
                        plusLength = Math.Round(plusLength / 0.25) * 0.25;
                        if (Math.Abs((plusLength + 1) * paper.width - w) > 0.25 * 0.1 * paper.width)
                        {
                            plusLength += 0.25;
                        }
                    }
                    ps = (PaperSize)Enum.Parse(typeof(PaperSize), paper.Size + "plus");

                }
            }
            catch (Exception)
            {
                ps = PaperSize.Undefined;
            }
            return ps;
        }
    }

    public enum PaperSize
    {
        /// <summary>
        /// The width or height of the page are set manually and override the PageSize property.
        /// </summary>
        Undefined = 0,


        /// <summary>
        /// Identifies a paper sheet size of 841 mm times 1189 mm or 33.11 inch times 46.81 inch.
        /// </summary>
        A0 = 1,

        /// <summary>
        /// Identifies a paper sheet size of 594 mm times 841 mm or 23.39 inch times 33.1 inch.
        /// </summary>
        A1 = 2,

        /// <summary>
        /// Identifies a paper sheet size of 420 mm times 594 mm or 16.54 inch times 23.29 inch.
        /// </summary>
        A2 = 3,

        /// <summary>
        /// Identifies a paper sheet size of 297 mm times 420 mm or 11.69 inch times 16.54 inch.
        /// </summary>
        A3 = 4,

        /// <summary>
        /// Identifies a paper sheet size of 210 mm times 297 mm or 8.27 inch times 11.69 inch.
        /// </summary>
        A4 = 5,

        /// <summary>
        /// Identifies a paper sheet size of 148 mm times 210 mm or 5.83 inch times 8.27 inch.
        /// </summary>
        A5 = 6,
        A0plus = 7,
        A1plus = 8,
        A2plus = 9,
        A3plus = 10,
        A4plus = 11,
        ////加长图框后面数值为加长的千分比，A0plus125，为A0+0.125
        //A0plus125 = 12,
        //A0plus250 = 13,
        //A0plus375 = 14,
        //A0plus500 = 15,
        //A0plus625 = 16,
        //A0plus750 = 17,
        //A0plus875 = 18,
        //A0plus1000 = 19,
        //A0plus1125 = 20,

    }
}
