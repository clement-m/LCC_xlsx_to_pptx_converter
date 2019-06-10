using System.IO;
using System.Collections.Generic;

namespace LCC_xlsx_to_pptx_converter.Class.datas
{
  class DeleteImages
  {
    public static void run(string pathFolder, List<string> listFile)
    {
      int fileNumber = 0;
      foreach(string fileName in listFile)
      {
        fileNumber++;
        Directory.Delete(pathFolder + "WorkBook" + fileNumber, true);
      }
    }
  }
}
