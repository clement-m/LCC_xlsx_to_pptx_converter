using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class convertTemplate
  {
    public static void run(A.Presentation template, A.Presentation newPresentation, int WB)
    {
      int slideId   = 0;
      int slidePres = 0;
      foreach (A.Slide slide in template.Slides)
      {
        cloneSlide.run(template, newPresentation, slideId);
        slideId++;
      }
    }
  }
}
