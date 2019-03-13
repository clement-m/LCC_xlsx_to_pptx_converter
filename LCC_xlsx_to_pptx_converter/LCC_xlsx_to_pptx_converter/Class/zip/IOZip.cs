using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Spreadsheet;


namespace LCC_xlsx_to_pptx_converter.Class
{
  class IOZip
  {

    const string DIR_PATH = @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\";

    // excel
    private void readCell(Cell cells, SharedStringTable sst)
    {
      // One way: go through each cell in the sheet
      foreach (Cell cell in cells)
      {
        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
        {
          int ssid = int.Parse(cell.CellValue.Text);
          string str = sst.ChildElements[ssid].InnerText;

          this.readCell(ssid, str);

          Console.WriteLine("Shared string {0}: {1}", ssid, str);
        }
        else if (cell.CellValue != null)
        {
          Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
        }
      }
    }
    private void readRow(Row rows, SharedStringTable sst)
    {
      // Or... via each row
      foreach (Row row in rows)
      {
        foreach (Cell c in row.Elements<Cell>())
        {
          if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
          {
            int ssid = int.Parse(c.CellValue.Text);
            string str = sst.ChildElements[ssid].InnerText;
            Console.WriteLine("Shared string {0}: {1}", ssid, str);
          }
          else if (c.CellValue != null)
          {
            Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
          }
        }
      }
    }
    private void readCell(int cellId, string value)
    {
      switch (cellId)
      {
        case 27:
          // RAPPORT INDOOR ORANGE
          break;
        case 29:
          // RAPPORT INDOOR ORANGE
          break;
        case 141:
          // VERSION CLIENT FINAL ?????
          break;
        case 142:
          // SERVICE VOIX
          break;
        case 143:
          // NIVEAU DU SIGNAL
          break;
        case 144:
          // qualité du signal
          break;
        case 154:
          // SERVICE DATA 3G
          break;
        case 4:
          // image
          break;
        default:
          // nothing ?
          break;
      }

    }

    // MANUAL SEARCH
    private void zipExtractMethod()
    {
      // zip unzip loadImage
      this.routine_zip_file();    //   I- create from file.xlsx => file.zip // FINISHED
      this.routine_unzip_xlsx();  //  II- unzip file.zip                    // FINISHED
      this.routine_get_imageFile();      // III- get text and image and store data // IN CREATION
    }
    private void routine_zip_file()
    {
      try
      {
        string[] xlsxList = Directory.GetFiles(DIR_PATH, "*.xlsx");

        foreach (string file in xlsxList)
        {
          string fileName = file.Substring(DIR_PATH.Length);

          string fileNewName = fileName.Substring(0, fileName.IndexOf('.')) + ".zip";

          try
          {
            System.IO.File.Copy(
              System.IO.Path.Combine(DIR_PATH, fileName),
              System.IO.Path.Combine(DIR_PATH, fileNewName)
            );

            Console.WriteLine("file " + fileName + ".xlsx converted in " + fileNewName + ".zip");
          }
          catch (IOException copyError)
          {
            Console.WriteLine(copyError.Message);
          }
        }
      }
      catch (DirectoryNotFoundException dirNotFound)
      {
        Console.WriteLine(dirNotFound.Message);
      }
    }
    private void routine_unzip_xlsx()
    {
      try
      {
        string[] zipList = Directory.GetFiles(DIR_PATH, "*.zip");

        foreach (string f in zipList)
        {
          string fName = f.Substring(DIR_PATH.Length);

          fName = fName.Substring(0, fName.IndexOf('.'));

          try
          {
            ZipFile.ExtractToDirectory(f, DIR_PATH + fName);
          }
          catch (Exception e)
          {
            Console.WriteLine(e.Message);
          }
        }

        foreach (string f in zipList)
        {
          System.IO.File.Delete(f);
        }
      }
      catch (DirectoryNotFoundException dirNotFound)
      {
        Console.WriteLine(dirNotFound.Message);
      }
    }
    private void routine_get_imageFile()
    {
      try
      {
        string xmlImageFolderPath = DIR_PATH + "Outdoor\\xl\\media\\";
        string[] imgList = Directory.GetFiles(xmlImageFolderPath, "*.*");

        foreach (string file in imgList)
        {
          string fileName = file.Substring(xmlImageFolderPath.Length);
          Image myImg = Image.FromFile(xmlImageFolderPath + fileName);

          XLSXData pptxdt = new XLSXData();
          //pptxdt.add(myImg);

          switch (fileName)
          {

          }
        }
      }
      catch (DirectoryNotFoundException dirNotFound)
      {
        Console.WriteLine(dirNotFound.Message);
      }
    }
  }
}
