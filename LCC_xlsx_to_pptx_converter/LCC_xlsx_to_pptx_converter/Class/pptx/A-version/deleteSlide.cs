using Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class deleteSlide
  {
    public static void Run(Presentation destPres, int slideNumber)
    {

      using (destPres)
      {
        destPres.Slides.RemoveAt(slideNumber);
      }
    }
  }
}
