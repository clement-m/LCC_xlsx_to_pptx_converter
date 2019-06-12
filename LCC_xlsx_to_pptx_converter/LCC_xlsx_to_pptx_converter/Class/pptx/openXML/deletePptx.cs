using System.IO;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.openXML
{
  class deletePptx
  {
    public static void run(string pptxFilePath)
    {
      if(File.Exists(pptxFilePath))
      {
        File.Delete(pptxFilePath);
      }
    }
  }
}
