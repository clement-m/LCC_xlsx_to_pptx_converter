using System.IO;

namespace LCC_xlsx_to_pptx_converter.Class.datas
{
  class DeleteImages
  {
    public static void run(string pathFolder)
    {
      Directory.Delete(pathFolder + "WorkBook1", true);
      Directory.Delete(pathFolder + "WorkBook2", true);
      Directory.Delete(pathFolder + "WorkBook3", true);
      Directory.Delete(pathFolder + "WorkBook4", true);
    }
  }
}
