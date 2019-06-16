using System.Collections.Generic;
using A = Aspose.Cells;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  public static class CopyXlsxFile
  {
    public static void run(List<string> listFile)
    {
      int nbTempFile = 0;

      foreach(string fileName in listFile)
      {
        nbTempFile++;

        A.Workbook workbook = new A.Workbook(fileName);

        workbook.Save(fileName + "_TEMP" + nbTempFile + ".xlsx", A.SaveFormat.Xlsx);
      }
    }
  }
}
