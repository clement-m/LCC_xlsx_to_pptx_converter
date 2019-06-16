using System.Collections.Generic;
using LCC_xlsx_to_pptx_converter.Class.A_version;
using LCC_xlsx_to_pptx_converter.Class.openXML;

namespace LCC_xlsx_to_pptx_converter.Class.main
{
  public static class convertProcess
  {
    public static void run(List<string> listFile, string title)
    {
      deletePptx.run(getProgramDirectory.run() + "\\pptx_template\\" + "NEW_TEMPLATE.pptx"); // DEBUG

      OpenXlsx.run(listFile);

      int tic;
      int slideId;

      createClientPart.run(listFile);

      convert.run(listFile);

      deleteImages.run(getProgramDirectory.run(), listFile);

      listFile.Clear();

      Clean.run(getProgramDirectory.run());
    }
  }
}