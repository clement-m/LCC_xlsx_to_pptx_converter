using System.Collections.Generic;
using Aspose.Cells;
using System;
using System.IO;
using LCC_xlsx_to_pptx_converter.Class.main;

namespace LCC_xlsx_to_pptx_converter.Class.xlsx
{
  public static class OpenXlsx
  {
    public static Data run(string dataDir)
    {
      Data D = new Data();

      List<Workbook> WorkBooks = new List<Workbook>();
      WorkBooks.Add(new Workbook(dataDir + "TEMP1.xlsx"));
      WorkBooks.Add(new Workbook(dataDir + "TEMP2.xlsx"));
      WorkBooks.Add(new Workbook(dataDir + "TEMP3.xlsx"));
      WorkBooks.Add(new Workbook(dataDir + "TEMP4.xlsx"));

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

          foreach (Aspose.Cells.Cell cell in worksheet.Cells)
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

          foreach (Aspose.Cells.Drawing.Picture pic in worksheet.Pictures)
          {
            imageNumber++;

            string fileName = dataDir +
              "\\WorkBook"  + workbookId  + "\\" +
              "\\WorkSheet" + worksheetId + "\\" + 
              imageNumber + ".png"
            ;

            if (!(Directory.Exists(dataDir + "\\WorkBook" + workbookId + "\\" + "\\WorkSheet" + worksheetId)))
            {
              Directory.CreateDirectory(dataDir + "\\WorkBook" + workbookId + "\\" + "\\WorkSheet" + worksheetId);
            }

            FileStream file = File.Create(fileName);

            byte[] data = pic.Data;

            file.Write(data, 0, data.Length);
          }
        }
      }

      return D;
    }
  }
}
