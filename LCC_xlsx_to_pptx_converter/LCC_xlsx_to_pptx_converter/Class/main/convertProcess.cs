﻿using System;
using System.Collections.Generic;
using System.Drawing;

using LCC_xlsx_to_pptx_converter.Class.xlsx.openXML;
using LCC_xlsx_to_pptx_converter.Class.xlsx.A_version;

using LCC_xlsx_to_pptx_converter.Class.pptx.openXML;
using LCC_xlsx_to_pptx_converter.Class.pptx.A_version;

using LCC_xlsx_to_pptx_converter.Class.datas;

using A = Aspose.Slides;

namespace LCC_xlsx_to_pptx_converter.Class.main
{
  public static class convertProcess
  {
    public static void run(List<string> listFile, string title)
    {
      deletePptx.run(getProgramDirectory.run() + "\\pptx_template\\" + "NEW_TEMPLATE.pptx");

      Console.WriteLine("LCC xlsx_to_pptx_convertor Launched...");
      Console.WriteLine("=====================================================");

      Console.WriteLine("Step II:\n-Opening and extract data");
      Data D = OpenXlsx.run(listFile);
      Console.WriteLine("done.");

      Console.WriteLine("Step II:\n-Create A-Version Presentation");
      
      int WB = 1;
      A.Presentation newPresentation;

      using (newPresentation = new A.Presentation())
      {
        using (A.Presentation template = new A.Presentation(
          getProgramDirectory.run() 
          + "\\pptx_template\\" 
          + "template.pptx"))
        {
          deleteSlide.Run(newPresentation, 0);

          int slideId = 0;
          foreach (A.Slide slide in template.Slides)
          {
            cloneSlide.run(template, newPresentation, slideId);
            slideId++;
          }


          slideId = 0;
          foreach (A.Slide slide in template.Slides)
          {
            int imageWanted;
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
                break;
              case 12:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 18);
                changeImageSlide.run2(newPresentation, imageWanted, WB, 19);
                break;
              case 13:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 20);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 21);
                break;
              case 14:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 22);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 23);
                break;
              case 15:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 24);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 25);
                break;
              case 16:
                // MERCI
                break;
              case 17:
                // NIVEAU XXX
                //changeImageSlide.run2(newPresentation, imageWanted, WB, 26); // LOGO ORANGE
                break;
              case 18:
                // SCANNER GSM...LTE
                break;
              case 19:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 27);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 28);
                break;
              case 20:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 29);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 30);
                break;
              case 21:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 31);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 32);
                break;
              case 22:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 33);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 34);
                break;
              case 23:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 35);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 36);
                break;
              case 24:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 36);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 37);
                break;
              case 25:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 38);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 39);
                break;
              case 26:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 40);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 41);
                break;
              case 27:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 42);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 43);
                break;
              case 28:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 44);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 45);
                break;
              case 29:
                // TITLE
                break;
              case 30:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 46);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 47);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 48);
                break;
              case 31:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 49);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 50);
                break;
              case 32:
                //changeImageSlide.run2(newPresentation, imageWanted, WB, 51); // LEGENDE
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 52);
                break;
              case 33:
                // TITLE
                break;
              case 34:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 53);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 54);
                break;
              case 35:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 55);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 56);
                break;
              case 36:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 57);
                break;
              case 37:
                // TITRRE
                break;
              case 38:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 58);
                break;
              case 39:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 59);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 60);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 61);
                break;
              case 40:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 62);
                break;
              case 41:
                // UN AUTRE TITRE
                break;
              case 42:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 63);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 64);
                break;
              case 43:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 65);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 66);
                break;
              case 44:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 67);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 68);
                break;
              case 45:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 69);
                break;
              case 46:
                // TITLE
                break;
              case 47:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 70);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 71);
                break;
              case 48:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 72); //d 1/2/3!!
                break;
              case 49:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 73);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 74);
                break;
              case 50:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 75);
                break;
              case 51:
                // UN TITRE QUI CHANGE PAS (NORMALEMENT) PAS VERIIF
                break;
              case 52:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 76);
                break;
              case 53:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 77);
                break;
              case 54:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 70);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 78);
                break;
              case 55:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 80);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 79);
                break;
              case 56:
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 82);
                imageWanted = 103;
                changeImageSlide.run2(newPresentation, imageWanted, WB, 81);
                break;
            }
            
            slideId++;
          }// end foreach
          
        }
        newPresentation.Save(getProgramDirectory.run() + "\\pptx_template\\" +  "NEW_TEMPLATE.pptx", A.Export.SaveFormat.Pptx);
      }

      Console.WriteLine("done.");

      Console.WriteLine("Step III:\n-Dispose data");
      D.dispose();
      Console.WriteLine("done.");

      Console.WriteLine("Step IV:\n-Cleaning Images");
      deleteImages.run(getProgramDirectory.run(), listFile);
      listFile.Clear();
      Console.WriteLine("done.");
      
      Console.WriteLine("Step V:\n-Converting A-Version to PPTX (with A)");
      Clean.run(getProgramDirectory.run());
      Console.WriteLine("-Convertion succeeded");

      newPresentation.Dispose();
      Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
      Console.ReadLine();
    }
  }
}
