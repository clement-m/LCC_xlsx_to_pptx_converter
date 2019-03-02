using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LCC_xlsx_to_pptx_converter.Class
{
  class xlsxClass
  {
    /*
        pptxData myData = new pptxData();

        try
        {
          string[] xlsxList = Directory.GetFiles(DIR_PATH, "*.xlsx");

          foreach (string f in xlsxList)
          {
            string fName = f.Substring(DIR_PATH.Length);

            if(fName == "Outdoor.xlsx") {
              Console.WriteLine("read " + fName);
              using (SpreadsheetDocument document = SpreadsheetDocument.Open(f, true))
              {
                WorkbookPart workbookPart = document.WorkbookPart;
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sst = sstpart.SharedStringTable;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet sheet = worksheetPart.Worksheet;

                myData.addImage(worksheetPart);

                pptxMaker pptxMaker = new pptxMaker(myData);
                OpenXmlUtils.CreatePresentation(DIR_PATH + "make.pptx");
                //pptxMaker.run();

                Console.WriteLine("end");

                //var cells = sheet.Descendants<Cell>();
                //var rows = sheet.Descendants<Row>();
                //this.readCell(cells, sst);
                //this.readRow(rows, sst);

                /*
                string roowww = "1";
                string col = "0";

                WorkbookPart wbPart = document.WorkbookPart;
                var workSheet = wbPart.WorksheetParts.FirstOrDefault();

                TwoCellAnchor cellHoldingPicture = workSheet.DrawingsPart.WorksheetDrawing.OfType<TwoCellAnchor>()
                     .Where(c => c.FromMarker.RowId.Text == roowww &&
                            c.FromMarker.ColumnId.Text == col).FirstOrDefault();

                var picture = cellHoldingPicture.OfType<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().FirstOrDefault();
                string rIdofPicture = picture.BlipFill.Blip.Embed;

                Console.WriteLine("The rID of this Anchor's [{0},{1}] Picture is '{2}'", roowww, col, rIdofPicture);

                *//*
              }
            }
          }
        }
        catch (Exception error)
        {
          Console.WriteLine(error.Message);
        }
        */
  }
}
