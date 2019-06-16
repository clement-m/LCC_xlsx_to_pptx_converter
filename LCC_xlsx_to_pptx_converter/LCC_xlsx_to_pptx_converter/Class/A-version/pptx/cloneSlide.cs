using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  class cloneSlide
  {
    public static void run(A.Presentation sourcePresentation, A.Presentation destPres,ref int switchTic, ref int slidePres)
    {
      using (sourcePresentation)
      {
        using (destPres)
        {
          A.ISlideCollection slideCollection = destPres.Slides;
          slideCollection.InsertClone(slidePres, sourcePresentation.Slides[switchTic]);
        }
      }
      slidePres++;
      switchTic++;
    }
  }
}