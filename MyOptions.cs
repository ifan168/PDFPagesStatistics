using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ifan.PDF;

namespace PDF页面统计
{
    [Serializable]
    public class MyOptions
    {
        public int tolerance { get; set; }
        public List<Paper> PaperSizes = new List<Paper>();
        public MyOptions()
        {

        }
    }
       
    public class OptionsHelper
    {

        string path = Environment.CurrentDirectory + "//PDFPaperSizeOptions.xml";
        public OptionsHelper()
        {

        }

        public MyOptions GetOptions()
        {
            try
            {
                return ReadFromXml(path);
            }
            catch (Exception)
            {
                return null;
            }
          
        }

        public void SetOptions(MyOptions myOptions)
        {

            writeToXml(myOptions);
        }
        private void writeToXml(MyOptions myOptions)
        {
            System.Xml.Serialization.XmlSerializer writer =
                               new System.Xml.Serialization.XmlSerializer(typeof(MyOptions));
            System.IO.FileStream file = System.IO.File.Create(path);
            writer.Serialize(file, myOptions);
            file.Close();
        }
        private MyOptions ReadFromXml(String path)
        {
            System.Xml.Serialization.XmlSerializer reader =
                  new System.Xml.Serialization.XmlSerializer(typeof(MyOptions));
            System.IO.StreamReader file = new System.IO.StreamReader(path);
            MyOptions result = new MyOptions();
            result = (MyOptions)reader.Deserialize(file);
            file.Close();
            return result;

        }
    }
}
