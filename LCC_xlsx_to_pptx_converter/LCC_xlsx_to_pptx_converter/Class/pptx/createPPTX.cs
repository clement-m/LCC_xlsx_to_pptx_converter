using Aspose.Slides;
using LCC_xlsx_to_pptx_converter.Class.main;
using System.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class.pptx
{
  public static class createPPTX
  {
    public static Aspose.Slides.Presentation run(string dataDir, Data D)
    {
      using (Presentation pres = new Presentation(dataDir + "BCNE_19S07_VDI_T-CSFB-D-S_OSM-CIMENTERIE-EQIOM.pptx"))
      {
        int slideId = 0;
        foreach (Slide slide in pres.Slides)
        {
          slideId++;

          switch (slideId)
          {
            case 11:
              //foreach (IShape shp in slide.Shapes)
              //if (shp.Placeholder != null)
              //{
                // Change the text of each placeholder
                //((IAutoShape)shp).TextFrame.Text = "This is Placeholder";
              //}
              break;
            case 12: // SERVICE VOIX Niveau RDC
              // remplacer image gauche par
              // workbook2 / worksheet1 / 60.png
              foreach (IShape shp in slide.Shapes)
              { 
                if (shp.Placeholder != null)
                {
                  if(!(shp is PictureFrame))
                  {
                    string text = ((IAutoShape)shp).TextFrame.Text;
                  } else {
                    int dza = 4561;
                    
                  }
                  

                    
                  string fileName = getImage(dataDir, 2, 1, 60); 

                  

                  // Set the picture
                  Image img = (Image)new Bitmap(fileName);
                  IPPImage imgx = pres.Images.AddImage(img);
                  shp.FillFormat.PictureFillFormat.Picture.Image = imgx;
                }
              }
              /*

              // Add autoshape of rectangle type
              IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150);


              // Set the fill type to Picture
              shp.FillFormat.FillType = FillType.Picture;

              // Set the picture fill mode
              shp.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Tile;

              // Set the picture
              System.Drawing.Image img = (System.Drawing.Image)new Bitmap(dataDir + "Tulips.jpg");
              IPPImage imgx = pres.Images.AddImage(img);
              shp.FillFormat.PictureFillFormat.Picture.Image = imgx;

              //Write the PPTX file to disk
              pres.Save(dataDir + "RectShpPic_out.pptx", SaveFormat.Pptx);

              */
              //ExEnd:FillShapesPicture
              break;
          }
        }

        return pres;
      }
    }

    public static string getImage(string dataDir, int WorkBook, int WorkSheet, int imageNumber)
    {
      return dataDir +
        "WorkBook" + WorkBook + "\\" +
        "WorkSheet" + WorkSheet + "\\" +
        imageNumber + ".png"
      ;
    }
  }
}
