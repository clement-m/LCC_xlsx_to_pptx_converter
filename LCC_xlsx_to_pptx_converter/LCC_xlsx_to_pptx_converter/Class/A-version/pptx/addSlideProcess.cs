using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  class addSlideProcess
  {
    public static void run(A.Presentation sourcePresentation, A.Presentation destPres, ref int switchTic, ref int slidePres, int WB,
      int imgIdSolo, int imageTargetSolo)
    {
      cloneSlide.run(sourcePresentation, destPres, ref switchTic, ref slidePres);

      changeImageSlide.run2(destPres, imgIdSolo, WB, imageTargetSolo);
    }

    public static void run(A.Presentation sourcePresentation, A.Presentation destPres, ref int switchTic, ref int slidePres, int WB,
      
      int imgIdLeft , int slideTarget,
      int imgIdRight, int slide2Target)
    {
      cloneSlide.run(sourcePresentation, destPres, ref switchTic, ref slidePres);

      changeImageSlide.run2(destPres, imgIdLeft, WB, slideTarget);
      changeImageSlide.run2(destPres, imgIdRight, WB, slide2Target);
    }

    public static void run(A.Presentation sourcePresentation, A.Presentation destPres, ref int switchTic, ref int slidePres, int WB,
       int imgIdLeft  , int slideTarget,
       int imgIdMiddle, int slide2Target,
       int imgIdRight , int slide3Target)
    {
      cloneSlide.run(sourcePresentation, destPres, ref switchTic, ref slidePres);

      changeImageSlide.run2(destPres, imgIdLeft, WB, slideTarget);
      changeImageSlide.run2(destPres, imgIdMiddle, WB, slide2Target);
      changeImageSlide.run2(destPres, imgIdRight, WB, slide3Target);
    }
  }
}
