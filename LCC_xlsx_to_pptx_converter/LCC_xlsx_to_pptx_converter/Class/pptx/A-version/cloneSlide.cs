using Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class cloneSlide
  {
    public static void run(Presentation sourcePresentation, Presentation destPres, int slideNumber)
    {
      using (sourcePresentation)
      {
        using (destPres)
        {
          ISlideCollection slideCollection = destPres.Slides;
          slideCollection.InsertClone(slideNumber, sourcePresentation.Slides[slideNumber]);
        }
      }
    }

    public static void run2(Presentation sourcePresentation, Presentation destPres, int slideNumber)
    {
      using (sourcePresentation)
      {
        using (destPres)
        {
          ISlide SourceSlide = sourcePresentation.Slides[slideNumber];
          IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

          IMasterSlideCollection masters = destPres.Masters;
          IMasterSlide DestMaster = SourceSlide.LayoutSlide.MasterSlide;

          IMasterSlide iSlide = masters.AddClone(SourceMaster);

          ISlideCollection slds = destPres.Slides;
          slds.AddClone(SourceSlide, iSlide, true);
        }
      }
    }
  }
}