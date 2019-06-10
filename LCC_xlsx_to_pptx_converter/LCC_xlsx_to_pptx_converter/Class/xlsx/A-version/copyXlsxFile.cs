using System.Collections.Generic;
using Aspose.Cells;
using LCC_xlsx_to_pptx_converter.Class.datas;

namespace LCC_xlsx_to_pptx_converter.Class.xlsx
{
  public static class CopyXlsxFile
  {
    public static void run(List<string> listFile)
    {
      int nbTempFile = 0;

      foreach(string fileName in listFile)
      {
        nbTempFile++;

        Workbook workbook = new Workbook(fileName);

        workbook.Save(fileName + "_TEMP" + nbTempFile + ".xlsx", SaveFormat.Xlsx);
      }
    }
  }
}
