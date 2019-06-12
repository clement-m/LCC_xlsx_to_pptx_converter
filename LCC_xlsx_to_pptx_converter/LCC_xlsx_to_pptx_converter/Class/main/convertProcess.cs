using System.Collections.Generic;
using LCC_xlsx_to_pptx_converter.Class.xlsx.A_version;
using LCC_xlsx_to_pptx_converter.Class.pptx.openXML;
using LCC_xlsx_to_pptx_converter.Class.pptx.A_version;
using LCC_xlsx_to_pptx_converter.Class.datas;
using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.main
{
  public static class convertProcess
  {
    public static void run(List<string> listFile, string title)
    {
      deletePptx.run(getProgramDirectory.run() + "\\pptx_template\\" + "NEW_TEMPLATE.pptx");

      Data D = OpenXlsx.run(listFile);

      int WB = 1;

      using (A.Presentation newPresentation = new A.Presentation())
      {

        using (A.Presentation template = new A.Presentation(
          getProgramDirectory.run() 
          + "\\pptx_template\\" 
          + "template.pptx"))
        {
          deleteSlide.run(newPresentation, 0);

          convertTemplate.run(template, newPresentation, WB);

          convertImages.run(template, newPresentation, WB);
        }

        newPresentation.Save(getProgramDirectory.run() + "\\pptx_template\\" +  "NEW_TEMPLATE.pptx", A.Export.SaveFormat.Pptx);
      }

      D.dispose();

      deleteImages.run(getProgramDirectory.run(), listFile);

      listFile.Clear();

      Clean.run(getProgramDirectory.run());
    }
  }
}