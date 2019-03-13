using System;
using System.Windows.Forms;
using LCC_xlsx_to_pptx_converter.Class;
using LCC_xlsx_to_pptx_converter.Class.pptx;
using System.Text;

using System.IO;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using System.Collections.Generic;



namespace LCC_xlsx_to_pptx_converter
{
  
  public partial class Form1 : Form
  {
    const string DIR_PATH = @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\";

    public Form1()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      PPTXData Data = new PPTXData();

      string path = DIR_PATH + "make.pptx";

      PPTXWriter.CreatePresentation(path);

      
      //PPTXWriter.createEmptySlide(path, "rId3", "rId4", false);

      int slideId = 1;

      //string templateName = "template";
      string templateName = "tt";

      string pathTemplate = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), @"..\\..\\pptx_template\\" + templateName + ".pptx");
      
      using (PresentationDocument presentationDocumentTemplate =
          PresentationDocument.Open(pathTemplate, false))
      {
        Console.WriteLine("Ouverture du template : {0}.pptx", templateName);

        // Get the relationship ID of the first slide.
        PresentationPart part = presentationDocumentTemplate.PresentationPart;
        //OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

        SlideIdList slideIdList = part.Presentation.SlideIdList;
        int ids = slideIdList.ChildElements.Count;

        Console.WriteLine("happy");

        for (int i = 0; i <= ids; i++)
        {
          if (i == 11)
          {
            int iString = i + 1;

            string relId                     = (slideIdList.ChildElements[i] as SlideId).RelationshipId;
            SlidePart slidePartTemplate      = (SlidePart)part.GetPartById(relId);
            List<string> imagesSlideTemplate = Data.getImages(slidePartTemplate, relId);

            ImagePart Img = (ImagePart)slidePartTemplate.ImageParts;

            //PPTXWriter.createEmptySlide(path, Img);

            Slide slide = PPTXWriter.createEmptySlide(path,
            @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\image1.png");

            PPTXWriter.InsertImageInLastSlide(path, @"D:\Xampp\htdocs\github\00-docs\extract analyze et rapport\image1.png");

            if (imagesSlideTemplate.Count == 0)
            {
              Console.WriteLine("Aucune image dans : Template tt.pptx - slide {0}", iString);
              /*
              using (PresentationDocument makePPTX = PresentationDocument.Open(path, true))
              {
                PPTXWriter.InsertNewSlide(makePPTX, i + 1, "KDOPAKDAZPKD");

                Console.WriteLine("clonage de: {0}", relId);
              }
              */
            } else {
              using (PresentationDocument makePPTX = PresentationDocument.Open(path, true))
              {
                //PPTXWriter.InsertNewSlideWithImage(slidePartTemplate, makePPTX, i, "KDOPAKDAZPKD", slidePartTemplate.ImageParts);

                //PPTXWriter.createEmptySlide(path);
                Console.WriteLine("clonage de: {0}", relId);
              }

              foreach (ImagePart imgInSlideTemplate in slidePartTemplate.ImageParts)
              {
                string imageNumber = PPTXReader.getNumberImage(imgInSlideTemplate);

                Console.WriteLine("slide {0} : image{1}.png", iString, imageNumber);

                Stream streamForImgTemplate = imgInSlideTemplate.GetStream();
                long length = streamForImgTemplate.Length;
                byte[] ImgConvertedToByte = new byte[length];
                //Data.addImagesData("image" + imageNumber + ".png", ImgConvertedToByte);


              }

              
              Console.WriteLine("dzadzadzad slide {0}", iString);


            }
            /*
            using (PresentationDocument makePPTX = PresentationDocument.Open(path, true))
            {
              PPTXWriter.InsertNewSlide(makePPTX, i + 1, "KDOPAKDAZPKD");

              Console.WriteLine("clonage de: {0}", relId);
            }
            */
            

            // Build a StringBuilder object.
            StringBuilder paragraphText = new StringBuilder();

            // Get the inner text of the slide.
            IEnumerable<Text> texts = slidePartTemplate.Slide.Descendants<Text>();
            foreach (Text text in texts)
            {
              //paragraphText.Append(text.Text);
              Console.WriteLine("text content {0}", text.Text);
              //switchSlide(relId, text.Text, slide);
            }
            string sldText = paragraphText.ToString();
          } // end if
        } // end for

      }









      Console.WriteLine("-add slide " + slideId + " to make.pptx");

      string imagePath = DIR_PATH + "image1.png";

      //Console.WriteLine("-add image " + imagePath + " to slide " + slideId + " in make.pptx");

      //PPTXClass.InsertImageInLastSlide(path, imagePath);

      slideId++;

      //Console.WriteLine("-add slide " + slideId + " to make.pptx");
      //PPTXWriter.addSlide(path, slideId, "test slide 3");
      
      //Console.WriteLine("-add image " + imagePath + " to slide " + slideId + " in make.pptx");
      //PPTXClass.InsertImageInLastSlide(imagePath);
      
      // SUITE DU PROGRAMME

      // Ajouter une image venant du fichier excel

      // verifier la correspondance

      // ajouter les deux images aux bonnes positions pour les gros titres

      // ajouter les images 

      
      Console.WriteLine("END PROGRAM");
      Console.WriteLine("CONVERTION XLSX to PPTX SUCCESS");
    }
  }
}