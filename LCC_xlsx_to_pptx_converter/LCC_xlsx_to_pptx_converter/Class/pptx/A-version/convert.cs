using System.Collections.Generic;
using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.pptx.A_version
{
  class convert
  {
    public static void run(A.Presentation template, A.Presentation newPresentation, List<string> listFile)
    {
      int slideId   = 0;
      int switchTic = 0;
      int WB        = 1;
      foreach (A.Slide slide in template.Slides)
      {
        switch (switchTic + 1)
        {
          case 0:
          case 1:
          case 2:
          case 3:
          case 4:
          case 5:
          case 6:
          case 7:
          case 8:
          case 9:
          case 10:
          case 11:
            cloneSlide.run(template, newPresentation, ref switchTic, ref slideId);
            break;
          case 12:
            addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
            103, 18, 
            104, 19);
            break;
          case 13:
            addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
              105, 20, 
              107, 21);
            break;
          case 14:
            addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
              106, 22, 
              108, 23);
            break;
          case 15:
            addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
              101, 24, 
              102, 25);

            ticChange.run(ref slideId, ref switchTic, ref WB, listFile);
            break;
          case 16: // MERCI
            cloneSlide.run(template, newPresentation, ref switchTic, ref slideId); 
            break;
        } // end switch
      } // end for
    } // end function
  } // end class
} // end namespace
