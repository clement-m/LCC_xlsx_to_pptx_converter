using System.IO;
using System.Collections.Generic;

namespace LCC_xlsx_to_pptx_converter.Class.datas
{
  class deleteImages
  {
    public static void run(string pathFolder, List<string> listFile)
    {
      int fileNumber = 0;
      if(listFile != null)
      {
        foreach (string fileName in listFile)
        {
          if (Directory.Exists(pathFolder + "WorkBook" + fileNumber))
          {
            fileNumber++;
            Directory.Delete(pathFolder + "WorkBook" + fileNumber, true);
          }
        }
      }
    }
  }
}
