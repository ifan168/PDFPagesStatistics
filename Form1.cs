using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ifan.PDF;
using System.Threading;
using System.Windows.Forms.VisualStyles;


namespace PDF页面统计
{
    public partial class Form1 : Form
    {
        private PdfDocument pdf1;
        private DialogResult dlgresult;
        List<string> lstpdfsUrl = new List<string>();
        List<FileInfo> fileInfos = new List<FileInfo>();

        /// <summary>
        /// 按页面储存PDF信息列表
        /// </summary>
        List<PDFPageStruct> MyPages = new List<PDFPageStruct>();


        public Form1()
        {
            InitializeComponent();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            dlgresult = openFileDialog1.ShowDialog();
            if (dlgresult == DialogResult.Cancel)
            {
                return;
            }
            lstpdfsUrl = openFileDialog1.FileNames.ToList();
            CountPDFpages(lstpdfsUrl);
        }

        private void CountPDFpages(List<string> lstPdfFiles)
        {
            CountPDFpages(lstpdfsUrl.ConvertAll(
                new Converter<string, FileInfo>(stringToFileInfo)));
        }

        private FileInfo stringToFileInfo(string input)
        {
            return new FileInfo(input);
        }

        private void GetPdfsSize(List<FileInfo> fileInfos)
        {

            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            MyPages.Clear();
            List<string> m_pdfssize = new List<string>();
            dataGridView1.Columns.Add("图纸名称", "图纸名称");
            foreach (FileInfo pdfFielInfo in fileInfos)
            {                
                pdf1 = PdfReader.Open(pdfFielInfo.FullName, PdfDocumentOpenMode.Import);
                PdfPages pages = pdf1.Pages;

                foreach (PdfPage item in pages)
                {

                    ifanPdfPaper ifanpdfpaper = new ifanPdfPaper(item);
                    MyPages.Add(new PDFPageStruct(pdfFielInfo.Directory.Name, pdfFielInfo, ifanpdfpaper));
                }
                pdf1.Close();
            }

            //根据有的图纸尺寸创建列
            var paparunits =
                  from p in MyPages
                  where p.m_pdfpaper.SizeName != PaperSize.Undefined.ToString()
                  orderby p.m_pdfpaper.PaperSizeHeight descending
                  group p by p.m_pdfpaper.SizeName;
            foreach (var item in paparunits)
            {
                dataGridView1.Columns.Add(item.Key, item.Key);
            }
            if (paparunits.Count() != MyPages.GroupBy(p => p.m_pdfpaper.SizeName).Count())
            {
                dataGridView1.Columns.Add(PaperSize.Undefined.ToString(), "未知尺寸");
                dataGridView1.Columns.Add("未知尺寸统计", "未知尺寸统计");
            }

            var groups =
                  from p in MyPages
                  orderby p.m_DirName ascending
                  group p by p.m_DirName;

            foreach (var item in groups)
            {

                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView1);
                row.Cells[0].Value = item.Key;

                var iis =
                    from p in item
                    group p by p.m_pdfpaper.SizeName;
                foreach (var tt in iis)
                {
                    row.Cells[dataGridView1.Columns[tt.Key].Index].Value = tt.Count();
                }
                //统计未定义尺寸图纸详细信息
                var undefpapers =
                        from p in item
                        where p.m_pdfpaper.SizeName == PaperSize.Undefined.ToString()
                        orderby p.m_pdfpaper.PaperSizeHeight descending
                        group p by (Math.Round(p.m_pdfpaper.PaperSizeWidth).ToString() + "x" + Math.Round(p.m_pdfpaper.PaperSizeHeight).ToString());
                foreach (var j in undefpapers)
                {
                    row.Cells[dataGridView1.Columns["未知尺寸统计"].Index].Value += string.Format("{0}:{1}\n", j.Key, j.Count());
                    foreach (var k in j)
                    {
                        PdfPages pages = k.m_pdfpaper.GetPdfPage().Owner.Pages;
                        int pageno = -1;
                        for (int i = 0; i < pages.Count; i++)
                        {
                            if (pages[i] == k.m_pdfpaper.GetPdfPage())
                            {
                                pageno = i + 1;
                            }
                        }
                        row.Cells[dataGridView1.Columns["未知尺寸统计"].Index].Value += string.Format("   {0}----第{1}页\n", k.m_FileInfo, pageno.ToString());
                    }
                }
                dataGridView1.Rows.Add(row);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            dlgresult = folderBrowserDialog1.ShowDialog();
            if (dlgresult == DialogResult.Cancel)
            {
                return;
            }
            string folderUrl = folderBrowserDialog1.SelectedPath;

            //string folderUrl = @"C:\Users\Administrator\Desktop\20200522简阳晒图\简阳土建蓝图PDF\给排水";
            List<FileInfo> lstfiles = GetAllFilesfromFolder(new DirectoryInfo(folderUrl));

            CountPDFpages(lstfiles);

        }

        private List<FileInfo> GetAllFilesfromFolder(DirectoryInfo DirectoryPath)
        {
            return GetFilesName(DirectoryPath);
        }

        private List<FileInfo> GetAllFilesFromFolder(string folderUrl)
        {
            DirectoryInfo root = new DirectoryInfo(folderUrl);
            List<FileInfo> lstfiles = GetFilesName(root);
            return lstfiles;
        }

        private void CountPDFpages(List<FileInfo> lstFileInfos)
        {
            if (lstFileInfos.Count > 0)
            {
                fileInfos.Clear();
                foreach (FileInfo pdfurl in lstFileInfos)
                {
                    fileInfos.Add(pdfurl);
                }
                GetPdfsSize(fileInfos);
            }
        }

        /// <summary>
        /// 获取文件夹下所有的PDF
        /// </summary>
        /// <param name="directoryInfo">需要获取的文件夹</param>
        /// <returns></returns>
        public List<FileInfo> GetFilesName(DirectoryInfo directoryInfo)
        {
            List<FileInfo> lstfiles = new List<FileInfo>();
            foreach (FileInfo item in directoryInfo.GetFiles())
            {
                if (item.Extension.ToLower().Contains("pdf"))
                {
                    lstfiles.Add(item);
                }
            }
            foreach (DirectoryInfo item in directoryInfo.GetDirectories())
            {
                lstfiles.AddRange(GetFilesName(item).ToArray());
            }
            return lstfiles;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("请在网页下方评论中填写意见，谢谢！");
            System.Diagnostics.Process.Start("http://www.bimzd.com/archives/2020/09/354");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.bimzd.com");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportExcel Et = new ExportExcel();
            Et.GridToExcel("PDF页面尺寸统计结果", dataGridView1);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // MessageBox.Show("注意:暂不支持一个PDF文件中含有两个及以上尺寸的页面!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dlgresult = folderBrowserDialog1.ShowDialog();
            if (dlgresult == DialogResult.Cancel)
            {
                return;
            }
            string folderUrl = folderBrowserDialog1.SelectedPath;
            DirectoryInfo root = new DirectoryInfo(folderUrl);

            if (MyPages.Count != 0)
            {
                //根据有的图纸尺寸创建列
                var paparunits =
                      from p in MyPages
                      orderby p.m_pdfpaper.PaperSizeHeight descending
                      group p by p.m_pdfpaper.SizeName;
                try
                {
                    //根据页面尺寸分类
                    foreach (var item in paparunits)
                    {
                        DirectoryInfo directory;
                        if (item.Key == PaperSize.Undefined.ToString())
                            directory = Directory.CreateDirectory(folderUrl + "\\" + "未知尺寸");
                        else
                            directory = Directory.CreateDirectory(folderUrl + "\\" + item.Key);

                        //根据文件名分类
                        foreach (var i in item.GroupBy(p => p.m_FileInfo))
                        {
                            //处理未知尺寸
                            if (item.Key == PaperSize.Undefined.ToString())
                            {

                                foreach (var j in i.GroupBy(p =>
                                Math.Round(p.m_pdfpaper.PaperSizeWidth).ToString() + "x" + Math.Round(p.m_pdfpaper.PaperSizeHeight).ToString()))
                                {
                                    directory = Directory.CreateDirectory(folderUrl + "\\" + "未知尺寸\\" + j.Key.ToString());
                                    //判断是否一个页面一个文件       

                                    if (i.Count() == MyPages.Where(p => p.m_FileInfo == i.Key).Count())
                                    {
                                        File.Copy(i.Key.FullName, directory.FullName + "\\" + i.Key.Name);
                                    }
                                    else
                                    {
                                        //一个文件包含多个页面
                                        foreach (var m in i.Where(p =>
                                        (Math.Round(p.m_pdfpaper.PaperSizeWidth).ToString() + "x" + Math.Round(p.m_pdfpaper.PaperSizeHeight).ToString()) == j.Key))
                                        {
                                            IEnumerable<PDFPageStruct> unPages = MyPages.Where(p => p.m_FileInfo == i.Key &&
                                            (Math.Round(p.m_pdfpaper.PaperSizeWidth).ToString() + "x" + Math.Round(p.m_pdfpaper.PaperSizeHeight).ToString()) == j.Key);
                                            //一个文件中所有页面都是同一尺寸
                                            if (MyPages.Where(p => p.m_FileInfo == i.Key).Count() == unPages.Count())
                                            {
                                                File.Copy(i.Key.FullName, directory.FullName + "\\" + i.Key.Name);
                                            }
                                            else
                                            {
                                                //一个文件中含有两种不同尺寸
                                                // 创建新的PDF文档
                                                PdfDocument document = new PdfDocument();

                                       
                                                foreach (var pdfpage in unPages.Select(p => p.m_pdfpaper.GetPdfPage()))
                                                {
                                                    document.AddPage(pdfpage);

                                                }
                                                // 保存文档

                                                document.Save(directory.FullName + "\\" + Path.GetFileNameWithoutExtension(i.Key.Name) + "-" + j.Key.ToString() + ".pdf");

                                            }

                                        }

                                    }
                                }
                            }
                            else
                            {
                                //判断是否一个页面一个文件                               
                                if (i.Count() == MyPages.Where(p => p.m_FileInfo == i.Key).Count())
                                {
                                    File.Copy(i.Key.FullName, directory.FullName + "\\" + i.Key.Name);
                                }
                                else
                                {
                                    //一个文件包含多个页面
                                    foreach (var m in i.GroupBy(p => p.m_pdfpaper.SizeName))
                                    {
                                        //一个文件中所有页面都是同一尺寸
                                        if (i.Count() == MyPages.Where(p => p.m_FileInfo == i.Key).Count())
                                        {
                                            File.Copy(i.Key.FullName, directory.FullName + "\\" + i.Key.Name);
                                        }
                                        else
                                        {
                                            //一个文件中含有两种不同尺寸
                                            //获取所有该文件下同尺寸页面
                                            var SameSizePdfs =
                                                from p in MyPages
                                                where p.m_FileInfo == i.Key && p.m_pdfpaper.SizeName == item.Key.ToString()
                                                select p.m_pdfpaper.GetPdfPage();
                                            // 创建新的PDF文档
                                            PdfDocument document = new PdfDocument();

                                            foreach (var pdfpage in SameSizePdfs)
                                            {
                                                document.AddPage(pdfpage);

                                            }
                                            // 保存文档

                                            document.Save(directory.FullName + "\\" + Path.GetFileNameWithoutExtension(i.Key.Name)  +"-" + m.Key + ".pdf");

                                        }
                                    }
                                }

                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                MessageBox.Show("创建成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                UpdataStatus(path);
                e.Effect = DragDropEffects.Link;
            }
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                List<FileInfo> m_lstfiles = new List<FileInfo>();
                string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop);
                UpdataStatus("就绪");
                foreach (string item in fileNames)
                {
                    FileInfo fi = new FileInfo(item);
                    FileAttributes fa = fi.Attributes;

                    if (fa == FileAttributes.Directory)
                    {
                        m_lstfiles.AddRange(GetAllFilesfromFolder(new DirectoryInfo(item)));
                    }
                    else if (fi.Extension.ToLower().Contains("pdf"))
                    {
                        m_lstfiles.Add(new FileInfo(item));
                    }
                }
                CountPDFpages(m_lstfiles);
            }


        }

        /// <summary>
        /// 更新状态栏label显示状况
        /// </summary>
        /// <param name="strText">显示的文字</param>
        private void UpdataStatus(string strText)
        {
            toolStripStatusLabel2.Text = strText;
        }

        private void dataGridView1_DragLeave(object sender, EventArgs e)
        {
            UpdataStatus("就绪");
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            FormOptions formOptions = new FormOptions();
            formOptions.ShowDialog();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MyOptions Options = new OptionsHelper().GetOptions();
            if (Options != null)
            {
                ifanPdfPaper.m_papers = Options.PaperSizes;
                ifanPdfPaper.intTolerance = Options.tolerance;
            }

        }
    }

    public class PDFPageStruct
    {
        public string m_DirName;
        public FileInfo m_FileInfo;
        public ifanPdfPaper m_pdfpaper;

        /// <summary>
        /// PDF图纸信息储存类
        /// </summary>
        /// <param name="dm">文件目录</param>
        /// <param name="fileInfo">文件FileInfo</param>
        /// <param name="m_pdfpaper">PDF页面</param>
        public PDFPageStruct(string DirName, FileInfo fileInfo, ifanPdfPaper ifanPdfPaper)
        {
            m_DirName = DirName;
            m_FileInfo = fileInfo;
            m_pdfpaper = ifanPdfPaper;
        }
    }

}


