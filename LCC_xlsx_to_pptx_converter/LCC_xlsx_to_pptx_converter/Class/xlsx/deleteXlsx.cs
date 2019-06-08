using System.IO;

namespace LCC_xlsx_to_pptx_converter.Class.xlsx
{
  class DeleteXlsx
  {
    public static void run(string path)
    {
      File.Delete(path + "TEMP1.xlsx");
      File.Delete(path + "TEMP2.xlsx");
      File.Delete(path + "TEMP3.xlsx");
      File.Delete(path + "TEMP4.xlsx");
    }
  }
}
