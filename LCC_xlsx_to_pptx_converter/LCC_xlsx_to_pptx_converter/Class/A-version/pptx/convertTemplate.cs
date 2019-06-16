using System.Collections.Generic;
using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  class convertTemplate
  {
    public static void run(A.Presentation template, A.Presentation newPresentation, List<string> listFile, int WB)
    {
      int slideId   = 0;
      int slidePres = 0;
      foreach (A.Slide slide in template.Slides)
      {
        switch (slideId + 1)
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
            //cloneSlide.run(template, newPresentation, slidePres);
            break;
          case 12:
          case 13:
          case 14:
          case 15:
            //slidePres = cloneListSlide.run(template, newPresentation, listFile, slidePres);
            break;
          case 16:
            //cloneSlide.run(template, newPresentation, slidePres); // MERCI
            break;
            /*
          case 17:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            // NIVEAU XXX
            //changeImageSlide.run2(newPresentation, imageWanted, WB, 26); // LOGO ORANGE
            break;
          case 18:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            // SCANNER GSM...LTE
            break;
          case 19:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 20:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 21:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 22:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 23:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 24:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 25:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 26:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 27:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 28:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 29:
            cloneSlide.run(template, newPresentation, slideId);// TITLE
            break;
          case 30:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 31:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 32:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 33:
            cloneSlide.run(template, newPresentation, slideId);// TITLE
            break;
          case 34:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 35:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 36:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 37:
            cloneSlide.run(template, newPresentation, slideId);// TITRRE
            break;
          case 38:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 39:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 40:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 41:
            cloneSlide.run(template, newPresentation, slideId);// UN AUTRE TITRE
            break;
          case 42:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 43:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 44:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 45:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 46:
            cloneSlide.run(template, newPresentation, slideId);// TITLE
            break;
          case 47:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 48:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 49:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 50:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 51:
            cloneSlide.run(template, newPresentation, slideId);// UN TITRE QUI CHANGE PAS (NORMALEMENT) PAS VERIIF
            break;
          case 52:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 53:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 54:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 55:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
          case 56:
            cloneListSlide.run(template, newPresentation, listFile, slideId);
            break;
            */
          } // switch

        slideId++;
      } // endfor
    } // function
  } // class
} // namespace
