using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  class deleteSlide
  {
    public static void run(A.Presentation destPres, int slideNumber)
    {
      using (destPres)
      {
        destPres.Slides.RemoveAt(slideNumber);
      }
    }
  }
}
