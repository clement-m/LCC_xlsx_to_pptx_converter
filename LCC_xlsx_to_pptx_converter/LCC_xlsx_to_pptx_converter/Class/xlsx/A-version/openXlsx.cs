using System.Collections.Generic;
using Aspose.Cells;
using A = Aspose.Cells.Drawing;
using System;
using System.IO;
using LCC_xlsx_to_pptx_converter.Class.main;
using LCC_xlsx_to_pptx_converter.Class.datas;

namespace LCC_xlsx_to_pptx_converter.Class.xlsx
{
  public static class OpenXlsx
  {
    public static Data run(List<string> listFile)
    {
      Data D = new Data();

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

          Console.WriteLine("WorkSheet N°" + worksheetId);

          foreach (Cell cell in worksheet.Cells)
          {
            int rowNumber    = cell.Row;
            int columnNumber = cell.Column;
            string textData  = cell.StringValue;

            Console.WriteLine("cell " + rowNumber + " " + columnNumber + "'s text data: " + textData);

            DataSet DS = new DataSet(
              workbookId,
              worksheetId,
              rowNumber,
              columnNumber,
              textData
            );

            D.addDataSet(DS);
          }

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

      return D;
    }
  }
}
