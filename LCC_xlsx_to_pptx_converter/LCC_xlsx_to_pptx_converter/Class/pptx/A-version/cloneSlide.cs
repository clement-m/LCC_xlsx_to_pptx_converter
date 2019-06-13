using Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class cloneSlide
  {
    public static void run(Presentation sourcePresentation, Presentation destPres,ref int switchTic, ref int slidePres)
    {
      using (sourcePresentation)
      {
        using (destPres)
        {
          ISlideCollection slideCollection = destPres.Slides;
          slideCollection.InsertClone(slidePres, sourcePresentation.Slides[switchTic]);
        }
      }
      slidePres++;
      switchTic++;
    }
  }
}