using System.Collections.Generic;
using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.A_version
{
  class convert
  {
    public static void run(List<string> listFile)
    {
      using (A.Presentation newPresentation = new A.Presentation())
      {
        using (A.Presentation template = new A.Presentation(
          getProgramDirectory.run()
          + "\\pptx_template\\"
          + "template.pptx"))
        {
          deleteSlide.run(newPresentation, 0);
          int slideId = 0;
          int switchTic = 0;
          int WB = 1;
          foreach (A.Slide slide in template.Slides)
          {
            switch (switchTic + 1)
            {
              case 17:
                int forcedSlideId = 0;
                switchTic++;
                cloneSlide.run(template, newPresentation, ref forcedSlideId, ref slideId);
                break;
              case 18: // NIVEAU XXX
                switchTic++;
                forcedSlideId = 0;
                cloneSlide.run(template, newPresentation, ref switchTic, ref slideId);
                //changeImageSlide.run2(newPresentation, imageWanted, WB, 26); // LOGO ORANGE
                break;
              case 19: // SCANNER GSM...LTE
                cloneSlide.run(template, newPresentation, ref switchTic, ref slideId);
                break;
              case 19:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                11, 27,
                12, 28);
                break;
              case 20:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                    23, 29,
                    14, 30);
                break;
              case 21:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  15, 31,
                  16, 32);
                break;
              case 22:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                    17, 33,
                    18, 34);
                break;
              case 23:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                    19, 35,
                    20, 36);
                break;
              case 24:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  22, 38,
                  21, 37);
                break;
              case 25:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  33, 38,
                  24, 39);
                break;
              case 26:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  25, 40,
                  26, 41);
                break;
              case 27:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  29, 42,
                  30, 43);
                break;
              case 28:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  34, 44,
                  27, 45);
                break;
              case 29:
                // TITLE
                break;
              case 30:

                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  1, 46, // FAUX
                  4, 47,
                  7, 48);
                break;
              case 31:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  5, 49,
                  8, 50);
                break;
              case 32:
                //changeImageSlide.run2(newPresentation, imageWanted, WB, 51); // LEGENDE
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  10, 52);
                break;
              case 33:
                // TITLE
                break;
              case 34:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  37, 53,
                  33, 54);
                break;
              case 35:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  60, 55,
                  61, 56);
                break;
              case 36:

                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  66, 57);
                break;
              case 37:
                // TITRRE
                break;
              case 38:
                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  73, 58);
                break;
              case 39:

                addSlideProcess.run(template, newPresentation, ref switchTic, ref slideId, WB,
                  63, 59,
                  58, 60,
                  74, 61);
                break;
              case 40:
                imageWanted = 76;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 62);
                break;
              case 41:
                // UN AUTRE TITRE
                break;
              case 42:
                imageWanted = 77;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 63);
                imageWanted = 78;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 64);
                break;
              case 43:
                imageWanted = 80;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 65);
                imageWanted = 81;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 66);
                break;
              case 44:
                imageWanted = 87;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 67);
                imageWanted = 88;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 68);
                break;
              case 45:
                imageWanted = 86;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 69);
                break;
              case 46:
                // TITLE
                break;
              case 47:
                imageWanted = 92;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 70); // mank 2
                imageWanted = 95;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 71);
                break;
              case 48:
                imageWanted = 96;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 72); //d 1/2/3!!
                break;
              case 49:
                imageWanted = 99;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 73);
                imageWanted = 100;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 74);
                break;
              case 50:
                imageWanted = 98;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 75);
                break;
              case 51:
                // UN TITRE QUI CHANGE PAS (NORMALEMENT) PAS VERIIF
                break;
              case 52:
                imageWanted = 92;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 76);
                break;
              case 53:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 77);
                break;
              case 54:
                imageWanted = 103;
                //changeImageSlide.run2(newPresentation, imageWanted, WB, 70);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 78);
                break;
              case 55:
                imageWanted = 53;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 80);
                imageWanted = 43;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 79);
                break;
              case 56:
                imageWanted = 54;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 82);
                imageWanted = 44;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 81);

                break;
            }
            slideId++;
          } // end switch
        }
        newPresentation.Save(getProgramDirectory.run() + "\\pptx_template\\" + "NEW_TEMPLATE.pptx", A.Export.SaveFormat.Pptx);
      }
    } // end function
  } // end class
} // end namespace
