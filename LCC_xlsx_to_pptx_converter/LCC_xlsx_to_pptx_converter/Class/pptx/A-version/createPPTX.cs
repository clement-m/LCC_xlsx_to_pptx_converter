using Aspose.Slides;
using LCC_xlsx_to_pptx_converter.Class.main;
using LCC_xlsx_to_pptx_converter.Class.pptx.A_version;
using LCC_xlsx_to_pptx_converter.Class.datas;

namespace LCC_xlsx_to_pptx_converter.Class.pptx
{
  public static class createPPTX
  {
    public static Presentation run(Data D, string customerName)
    {
      using (Presentation pres = new Presentation(getProgramDirectory.run() + "BCNE_19S07_VDI_T-CSFB-D-S_OSM-CIMENTERIE-EQIOM.pptx"))
      {
        int slideId = 0;
        foreach (Slide slide in pres.Slides)
        {
          int imageWanted;
          int slideTargetImage;
          
          slideId++;

          int WB = 1;
          switch (slideId)
          {
            case 11:
              changeTextInSlide.run(slide, customerName);
              break;
            case 12:
              imageWanted = 103;
              slideTargetImage = 12;
              changeImageSlide.run(pres, WB, imageWanted, slideTargetImage);
              break;
            case 126:
              imageWanted = 11;
              slideTargetImage = 59;
              changeImageSlide.run(pres, WB, imageWanted, slideTargetImage);
              imageWanted = 11;
              slideTargetImage = 60;
              changeImageSlide.run(pres, WB, imageWanted, slideTargetImage);
              break;
          }
        }
        return pres;
      }
    }
  }
}
