using A = Aspose.Slides;
using LCC_xlsx_to_pptx_converter.Class.datas;
using System.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class templatize
  {
    public static void run(int WB)
    {
      using (A.Presentation template =
        new A.Presentation(
          getProgramDirectory.run() + "/rap1/" + "BCNE_19S12_VDI_T-CSFB-D-S_SOCOTEC-MULHOUSE.pptx"))
      {
        int imageWanted;
        int slideTargetImage;

        imageWanted = 12;
        slideTargetImage = 0;

        Image newImage = Image.FromFile(getImage.run(getProgramDirectory.run(), WB, imageWanted));
        for (int y = 1; y <= template.Images.Count - 1; y++)
        {
          changeImageSlide.run(template, imageWanted, WB, y);
        }

        newImage.Dispose();

        foreach (A.Slide slide in template.Slides)
        {
          changeTextInSlide.run(slide, "");

          template.Save(
            getProgramDirectory.run() + "TEMPLATE.pptx",
            Aspose.Slides.Export.SaveFormat.Pptx
          );
          template.Dispose();
        }
      }
    }
  }
}
