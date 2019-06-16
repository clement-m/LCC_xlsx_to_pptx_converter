using System.Collections.Generic;
using System.IO;
using LCC_xlsx_to_pptx_converter.Class.openXML;
using Aspose.Cells;
using A = Aspose.Cells.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  public static class OpenXlsx
  {
    public static void run(List<string> listFile)
    {
      string dataDir = getProgramDirectory.run();

      List<Workbook> WorkBooks = new List<Workbook>();

      int fileNumber = 0;
      foreach(string fileName in listFile)
      {
        fileNumber++;
        WorkBooks.Add(new Workbook(fileName));
      }

      int workbookId = 0;
      foreach(Workbook workbook in WorkBooks)
      {
        workbookId++;
        int worksheetId = 0;
        int imageNumber = 0;
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
          worksheetId++;

          foreach (A.Picture pic in worksheet.Pictures)
          {
            imageNumber++;

            string fileName = dataDir +
              "\\WorkBook"  + workbookId  + "\\" +
              imageNumber + ".png"
            ;

            if (!(Directory.Exists(dataDir + "\\WorkBook" + workbookId)))
            {
              Directory.CreateDirectory(dataDir + "\\WorkBook" + workbookId);
            }

            FileStream file = File.Create(fileName);

            byte[] data = pic.Data;

            file.Write(data, 0, data.Length);

            file.Close();

            file.Dispose();
          }
        }
      }
    }
  }
}
