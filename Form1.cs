using System;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Xml;

using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace LCC_xlsx_to_pptx_converter
{
  public partial class Form1 : Form
  {
    const string DIR_PATH = @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\";
    const string DIR_TEXT_PATH = @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\unzipped\xl\";
    const string DIR_MEDIA_PATH = @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\unzipped\xl\media\";

    public Form1()
    {
      InitializeComponent();
    }

    private void readXlsx()
    {


    }

    private void routine_zip_file()
    {
      try
      {
        string[] xlsxList = Directory.GetFiles(DIR_PATH, "*.xlsx");

        foreach(string f in xlsxList)
        {
          string fName = f.Substring(DIR_PATH.Length);

          string fNewName = fName.Substring(0, fName.IndexOf('.')) + ".zip";

          try
          {
            System.IO.File.Copy(Path.Combine(DIR_PATH, fName), Path.Combine(DIR_PATH, fNewName));
          }
          catch(IOException copyError)
          {
            Console.WriteLine(copyError.Message);
          }
        }
      }
      catch(DirectoryNotFoundException dirNotFound)
      {
        Console.WriteLine(dirNotFound.Message);
      }
    }
    private void routine_unzip_xlsx()
    {
      try
      {
        string[] zipList = Directory.GetFiles(DIR_PATH, "*.zip");

        foreach(string f in zipList)
        {
          string fName = f.Substring(DIR_PATH.Length);

          fName = fName.Substring(0, fName.IndexOf('.'));

          try
          {
            ZipFile.ExtractToDirectory(f, DIR_PATH + fName);
          } catch(Exception e) {
            Console.WriteLine(e.Message);
          }
        }

        foreach(string f in zipList)
        {
          System.IO.File.Delete(f);
        }
      }
      catch(DirectoryNotFoundException dirNotFound)
      {
        Console.WriteLine(dirNotFound.Message);
      }
    }

    private void routine_get_data()
    {
      try
      {
        string[] textList = Directory.GetFiles(DIR_TEXT_PATH, "sharedStrings.xml");

        string[] imgList = Directory.GetFiles(DIR_MEDIA_PATH, "*.*");

        foreach (string f in textList)
        {
          XmlDocument doc = new XmlDocument(); 
          doc.Load(f);

          XmlNode node = doc.DocumentElement.SelectSingleNode("/book/title");
          foreach (XmlNode node2 in doc.DocumentElement.ChildNodes)
          {
            string text = node2.InnerText; //or loop through its children as well
          }
        }
      }
      catch (DirectoryNotFoundException dirNotFound)
      {
        Console.WriteLine(dirNotFound.Message);
      }
    }
    
    private void button1_Click(object sender, EventArgs e)
    {
      object test = sender;
      EventArgs eventtt = e;

      string sSelectedFile = "";
      OpenFileDialog choofdlog = new OpenFileDialog();
      choofdlog.Filter = "All Files (*.*)|*.*";
      choofdlog.FilterIndex = 1;
      choofdlog.Multiselect = true;

      //FUNCTIONS                   //   STEP                                 // STATUS      

      this.routine_zip_file();    //   I- create from file.xlsx => file.zip // FINISHED
      this.routine_unzip_xlsx();  //  II- unzip file.zip                    // FINISHED
      this.routine_get_data();      // III- get text and image and store data // IN CREATION

      //this.routine_create_pptx(); //  IV- make pptx using data from xlsx    // IN QUEUE
    }

    /*
    private void routine_create_pptx()
    {
      // forEach slide
      this.routine_create_pptx_slide();
    }
    */
  }
}
