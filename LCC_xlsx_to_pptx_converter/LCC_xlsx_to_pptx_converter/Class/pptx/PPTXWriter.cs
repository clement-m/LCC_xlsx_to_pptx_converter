using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using F       = System.IO;
using P       = DocumentFormat.OpenXml.Presentation;
using D       = DocumentFormat.OpenXml.Drawing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using PIC     = DocumentFormat.OpenXml.Drawing.Pictures;

using OFFICE2010 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class
{
  public static class PPTXWriter
  {
    public static void CreatePresentation(string path)
    {
      int position = 0;
      string layoutName = "test name layout 1";

      if (System.IO.File.Exists(path))
      {
        System.IO.File.Delete(path);
        Console.WriteLine("DEBUG --- delete make.pptx for remake a new pptx");
      }
      PresentationDocument presentationDoc = PresentationDocument.Create(
          path,
          PresentationDocumentType.Presentation
      );
      PresentationPart presentationPart = presentationDoc.AddPresentationPart();
      presentationPart.Presentation = new Presentation();

      CreatePresentationParts(presentationPart);

      //addSlide(presentationPart, "rId3", "rId4");
      //addSlide(presentationPart, "rId6", "rId7");
      //addSlide(presentationPart, "rId9", "rId8");

      presentationDoc.Close();
    }

    public static void addSlide(string path, int position, string titleSlide)
    { 
      using (PresentationDocument presentationDoc = PresentationDocument.Open(path, true))
      {
        // Insert other code here.
        InsertNewSlide(presentationDoc, position, titleSlide);
      }
    }

    // Specify the non-visual properties of the new slide.
    public static void constructSlideContent(Slide slide)
    {
      P.NonVisualGroupShapeProperties nonVisualProperties = 
        slide.CommonSlideData.ShapeTree.AppendChild(new P.NonVisualGroupShapeProperties());

      nonVisualProperties.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() {
        Id = 1,
        Name = ""
      };
      
      nonVisualProperties.NonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
      
      nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();
    }
    // Specify the required shape properties for the title shape. 
    public static void specifyProtertyTitleShape(P.Shape titleShape, uint drawingObjectId)
    {
      titleShape.NonVisualShapeProperties =
        new P.NonVisualShapeProperties(
          new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
          new P.NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
          new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })
        )
      ;
      titleShape.ShapeProperties = new P.ShapeProperties();
    }
    // Specify the text of the title shape.
    public static void specifyTextTitleShape(P.Shape titleShape, string titleText)
    {
      titleShape.TextBody =
        new P.TextBody(
          new Drawing.BodyProperties(),
          new Drawing.ListStyle(),
          new Drawing.Paragraph(
            new Drawing.Run(
              new Drawing.Text() {
                Text = titleText
              }
            )
          )
        )
      ;
    }
    // Specify the required shape properties for the body shape.
    public static void specifyPropBodyShape(P.Shape bodyShape, uint drawingObjectId)
    {
      bodyShape.NonVisualShapeProperties = 
        new P.NonVisualShapeProperties(
          new P.NonVisualDrawingProperties() 
          { 
            Id = drawingObjectId, Name = "Content Placeholder"
          },
          new P.NonVisualShapeDrawingProperties(
              new Drawing.ShapeLocks()
              {
                NoGrouping = true
              }
          ),
          new ApplicationNonVisualDrawingProperties(
            new PlaceholderShape() 
            {
              Index = 1
            }
          )
        )
      ;
      bodyShape.ShapeProperties = new P.ShapeProperties();
    }
    // Specify the text of the body shape.
    public static void specifyTextBodyShape(P.Shape bodyShape, string text)
    {
      bodyShape.TextBody = 
        new P.TextBody(
          new Drawing.BodyProperties(),
          new Drawing.ListStyle(),
          new Drawing.Paragraph()
        )
      ;
    }

    /// <summary>
    /// Insert Image into Slide
    /// </summary>
    /// <param name="filePath">PowerPoint Path</param>
    /// <param name="imagePath">Image Path</param>
    /// <param name="imageExt">Image Extension</param>
    public static void InsertImageInLastSlide(string file, string imagePath)
    {
      using (var presentation = PresentationDocument.Open(file, true))
      {
        // Get the ID of the previous slide.
        SlidePart lastSlidePart;
        if (prevSlideId != null)
        {
          lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
        }
        else
        {
          lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
        }


        var slidePart = presentation
            .PresentationPart
            .SlideParts
            .First();

        var part = slidePart
            .AddImagePart(ImagePartType.Png);

        using (var stream = File.OpenRead(image))
        {
          part.FeedData(stream);
        }

        var tree = slidePart
            .Slide
            .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
            .First();

        var picture = new DocumentFormat.OpenXml.Presentation.Picture();

        picture.NonVisualPictureProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties();
        picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
        {
          Name = "My Shape",
          Id = (UInt32)tree.ChildElements.Count - 1
        });

        var nonVisualPictureDrawingProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties();
        nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
        {
          NoChangeAspect = true
        });
        picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

        var blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
        var blip1 = new DocumentFormat.OpenXml.Drawing.Blip()
        {
          Embed = slidePart.GetIdOfPart(part)
        };
        var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
        var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
        {
          Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
        };
        var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
        {
          Val = false
        };
        useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
        blipExtension1.Append(useLocalDpi1);
        blipExtensionList1.Append(blipExtension1);
        blip1.Append(blipExtensionList1);
        var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
        stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
        blipFill.Append(blip1);
        blipFill.Append(stretch);
        picture.Append(blipFill);

        picture.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
        picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
        {
          X = 0,
          Y = 0,
        });
        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
        {
          Cx = 1000000,
          Cy = 1000000,
        });
        picture.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
        {
          Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
        });

        tree.Append(picture);
      }
    }

    public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
    {
      if (presentationDocument == null) throw new ArgumentNullException("presentationDocument");
      if (slideTitle == null) throw new ArgumentNullException("slideTitle");
      if (presentationDocument.PresentationPart == null) throw new InvalidOperationException("The presentation document is empty.");

      PresentationPart presentationPart = presentationDocument.PresentationPart;
      Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
      uint drawingObjectId = 1;

      constructSlideContent(slide);

      // Specify the group shape properties of the new slide.
      slide.CommonSlideData.ShapeTree.AppendChild(new P.GroupShapeProperties());

      drawingObjectId++;

      // Declare and instantiate the title shape of the new slide.
      P.Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
      specifyProtertyTitleShape(titleShape, drawingObjectId);
      specifyTextTitleShape(titleShape, slideTitle);

      drawingObjectId++;

      // Declare and instantiate the body shape of the new slide.
      P.Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
      specifyPropBodyShape(bodyShape, drawingObjectId);
      specifyTextBodyShape(bodyShape, "blabla");

      // Create the slide part for the new slide.
      SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

      // Save the new slide part.
      slide.Save(slidePart);

      // Modify the slide ID list in the presentation part.
      // The slide ID list should not be null.
      SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
      /*
      uint maxSlideId = getMaxSlideId(slideIdList);
      SlideId prevSlideId = getPrevSlideId(slideIdList, position);

      // Get the ID of the previous slide.
      SlidePart lastSlidePart;

      if (prevSlideId != null)
      {
        lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
      }
      else
      {
        lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
      }

      // Use the same slide layout as that of the previous slide.
      if (null != lastSlidePart.SlideLayoutPart)
      {
        slidePart.AddPart(lastSlidePart.SlideLayoutPart);
      }



      // Insert the new slide into the slide list after the previous slide.
      SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
      newSlideId.Id = maxSlideId;
      newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

      // Save the modified presentation.
      presentationPart.Presentation.Save();
      */
    }

    public static void insertSlideAtId(
      PresentationDocument presentationDocument, 
      int position, 
      string slideTitle)
    {
      if (presentationDocument == null) throw new ArgumentNullException("presentationDocument");
      if (slideTitle == null) throw new ArgumentNullException("slideTitle");
      if (presentationDocument.PresentationPart == null) throw new InvalidOperationException("The presentation document is empty.");
      
      PresentationPart presentationPart = presentationDocument.PresentationPart;
      Slide            slide            = new Slide(new CommonSlideData(new ShapeTree()));
      uint drawingObjectId = 1;

      constructSlideContent(slide);

      // Specify the group shape properties of the new slide.
      slide.CommonSlideData.ShapeTree.AppendChild(new P.GroupShapeProperties());

      drawingObjectId++;

      // Declare and instantiate the title shape of the new slide.
      P.Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
      specifyProtertyTitleShape(titleShape, drawingObjectId);
      specifyTextTitleShape(titleShape, slideTitle);

      drawingObjectId++;

      // Declare and instantiate the body shape of the new slide.
      P.Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
      specifyPropBodyShape(bodyShape, drawingObjectId);
      specifyTextBodyShape(bodyShape, "blabla");

      // Create the slide part for the new slide.
      SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

      // Save the new slide part.
      slide.Save(slidePart);

      // Modify the slide ID list in the presentation part.
      // The slide ID list should not be null.
      SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
      /*
      uint maxSlideId = getMaxSlideId(slideIdList);
      SlideId prevSlideId = getPrevSlideId(slideIdList, position);

      // Get the ID of the previous slide.
      SlidePart lastSlidePart;

      if (prevSlideId != null)
      {
        lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
      }
      else
      {
        lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
      }

      // Use the same slide layout as that of the previous slide.
      if (null != lastSlidePart.SlideLayoutPart)
      {
        slidePart.AddPart(lastSlidePart.SlideLayoutPart);
      }



      // Insert the new slide into the slide list after the previous slide.
      SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
      newSlideId.Id = maxSlideId;
      newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

      // Save the modified presentation.
      presentationPart.Presentation.Save();
      */
    }

    public static void InsertNewSlideWithImage(
      SlidePart slidePartTemplate,
      PresentationDocument presentationDocument, 
      int position, 
      string slideTitle,
      IEnumerable<ImagePart> imgPart
      )
    {
      if (presentationDocument == null) throw new ArgumentNullException("presentationDocument");
      if (slideTitle == null) throw new ArgumentNullException("slideTitle");
      if (presentationDocument.PresentationPart == null) throw new InvalidOperationException("The presentation document is empty.");

      PresentationPart presentationPart = presentationDocument.PresentationPart;
      SlidePart        slidePart        = presentationPart.AddNewPart<SlidePart>();
      Slide            slide            = new Slide(new CommonSlideData(new ShapeTree()));
      P.Shape          titleShape       = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
      P.Shape          bodyShape        = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
      SlideIdList      slideIdList      = (presentationPart.Presentation.SlideIdList == null)
                                            ? presentationPart.Presentation.SlideIdList = new SlideIdList()
                                            : presentationPart.Presentation.SlideIdList
                                          ;

      uint drawingObjectId = 1;

      constructSlideContent(slide);

      // Specify the group shape properties of the new slide.
      slide.CommonSlideData.ShapeTree.AppendChild(new P.GroupShapeProperties());

      drawingObjectId++;

      // Declare and instantiate the title shape of the new slide.
      specifyProtertyTitleShape(titleShape, drawingObjectId);
      specifyTextTitleShape(titleShape, slideTitle);
      drawingObjectId++;

      // Declare and instantiate the body shape of the new slide.
      specifyPropBodyShape(bodyShape, drawingObjectId);
      specifyTextBodyShape(bodyShape, "blabla");

      //slidePart.ImageParts = imgPart;
      foreach (ImagePart image in imgPart)
      {
        ImagePart imageClone = slidePart.AddImagePart(image.ContentType, slidePartTemplate.GetIdOfPart(image));
        using (var imageStream = image.GetStream())
        {
          imageClone.FeedData(imageStream);
        }
      }
      // Save the new slide part.
      slide.Save(slidePart);

      uint maxSlideId = 1;
      foreach (SlideId slideId in slideIdList.ChildElements)
      {
        if (slideId.Id > maxSlideId)
        {
          maxSlideId = slideId.Id;
        }
      }
      maxSlideId++;

      SlideId prevSlideId = null;

      if(position == 11) 
      {
        // Get the ID of the previous slide.
        SlidePart lastSlidePart = null;
        if (prevSlideId != null)
        {
          lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
        }
        else
        {
          //lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
        }


      } else {
        slidePart.AddPart(slidePart.SlideLayoutPart);
      }


      // Insert the new slide into the slide list after the previous slide.
      //SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
      SlideId newSlideId = slideIdList.InsertAt(new SlideId(), 0);
      newSlideId.Id = maxSlideId;
      newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

      // Save the modified presentation.
      presentationPart.Presentation.Save();
    }

    public static SlidePart clone(ref PresentationPart presentationPart, SlidePart slideTemplate)
    {
      // Clone slide contents
      SlidePart slidePartClone = presentationPart.AddNewPart<SlidePart>();
      using (var templateStream = slideTemplate.GetStream(F.FileMode.Open))
      {
        slidePartClone.FeedData(templateStream);
      }

      // Copy layout part
      slidePartClone.AddPart(slideTemplate.SlideLayoutPart);

      // Copy the image parts
      foreach (ImagePart image in slideTemplate.ImageParts)
      {
        ImagePart imageClone = slidePartClone.AddImagePart(image.ContentType, slideTemplate.GetIdOfPart(image));
        using (var imageStream = image.GetStream())
        {
          imageClone.FeedData(imageStream);
        }
      }



      return slidePartClone;
    }



    public static Slide createEmptySlide(
      string path, 
      string imagePath = null,
      string masterId = "rId1", 
      string slideId = "rId2",
      bool isFirstSlide = true)
    {
      SlidePart slidePart1;
      SlideLayoutPart slideLayoutPart1;
      SlideMasterPart slideMasterPart1;
      ThemePart themePart1;

      using (PresentationDocument presentationDocument =
        PresentationDocument.Open(path, true))
      {
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        slidePart1 = CreateSlidePart(presentationPart, slideId);
        Slide slide = slidePart1.Slide;
        slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
        slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);

        slideMasterPart1.AddPart(slideLayoutPart1, masterId);
        presentationPart.AddPart(slideMasterPart1, masterId);

        

        if (isFirstSlide)
        {
          themePart1 = CreateTheme(slideMasterPart1);
          presentationPart.AddPart(themePart1, "rId5");
        }

        return slide;
      }
    }

    private static void CreatePresentationParts(PresentationPart presentationPart, string masterId = "rId1", string slideId = "rId2")
    {
      SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = masterId });
      SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = slideId });
      SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
      NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
      DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

      presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);
    }
    private static SlidePart CreateSlidePart(PresentationPart presentationPart, string slideId = "rId2")
    {
      SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>(slideId);
      Slide slide = null;
      
      slidePart1.Slide = 
      new Slide(
        new CommonSlideData(
          new ShapeTree(
            new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()
            ),
            new P.GroupShapeProperties(
              new TransformGroup()
            ),
            new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(
                  new PlaceholderShape()
                )
              ),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(
                  new EndParagraphRunProperties() { Language = "en-US" }
                )
              )
            )
          )
        ),
        new ColorMapOverride(
          new MasterColorMapping()
        )
      );

      return slidePart1;
    }
    private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1, string masterId = "rId1")
    {
      SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>(masterId);
      SlideLayout slideLayout = 
      new SlideLayout(
        new CommonSlideData(
          new ShapeTree(
            new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties()
              {
                Id = (UInt32Value)1U, Name = ""
              },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()
            ),
            new P.GroupShapeProperties(
              new TransformGroup()
            ),
            new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(
                  new ShapeLocks() { NoGrouping = true }
                ),
                new ApplicationNonVisualDrawingProperties(
                  new PlaceholderShape()
                )
              ),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(
                  new EndParagraphRunProperties()
                )
              )
            )
          )
        ),
        new ColorMapOverride(
          new MasterColorMapping()
        )
      );

      slideLayoutPart1.SlideLayout = slideLayout;

      return slideLayoutPart1;
    }
    private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1, string masterId = "rId1")
    {
      SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>(masterId);
      SlideMaster slideMaster = 
      new SlideMaster(
        new CommonSlideData(
          new ShapeTree(
            new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()
            ),
            new P.GroupShapeProperties(
              new TransformGroup()
            ),
            new P.Shape(
              new P.NonVisualShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
              new P.NonVisualShapeDrawingProperties(
              new ShapeLocks() { NoGrouping = true }
            ),
            new ApplicationNonVisualDrawingProperties(
              new PlaceholderShape() { Type = PlaceholderValues.Title }
            )
          ),
          new P.ShapeProperties(),
          new P.TextBody(
            new BodyProperties(),
            new ListStyle(),
            new Paragraph()
              )
            )
          )
        ),
        new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
        new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
        new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()
      ));

      slideMasterPart1.SlideMaster = slideMaster;

      return slideMasterPart1;
    }



















    private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
    {
      ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
      D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

      D.ThemeElements themeElements1 = new D.ThemeElements(
      new D.ColorScheme(
        new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
        new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
        new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
        new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
        new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
        new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
        new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
        new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
        new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
        new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
        new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
        new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
      { Name = "Office" },
        new D.FontScheme(
        new D.MajorFont(
        new D.LatinFont() { Typeface = "Calibri" },
        new D.EastAsianFont() { Typeface = "" },
        new D.ComplexScriptFont() { Typeface = "" }),
        new D.MinorFont(
        new D.LatinFont() { Typeface = "Calibri" },
        new D.EastAsianFont() { Typeface = "" },
        new D.ComplexScriptFont() { Typeface = "" }))
        { Name = "Office" },
        new D.FormatScheme(
        new D.FillStyleList(
        new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
        new D.GradientFill(
          new D.GradientStopList(
          new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
            new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
          { Position = 0 },
          new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
           new D.SaturationModulation() { Val = 300000 })
          { Val = D.SchemeColorValues.PhColor })
          { Position = 35000 },
          new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
           new D.SaturationModulation() { Val = 350000 })
          { Val = D.SchemeColorValues.PhColor })
          { Position = 100000 }
          ),
          new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
        new D.NoFill(),
        new D.PatternFill(),
        new D.GroupFill()),
        new D.LineStyleList(
        new D.Outline(
          new D.SolidFill(
          new D.SchemeColor(
            new D.Shade() { Val = 95000 },
            new D.SaturationModulation() { Val = 105000 })
          { Val = D.SchemeColorValues.PhColor }),
          new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
        {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
        },
        new D.Outline(
          new D.SolidFill(
          new D.SchemeColor(
            new D.Shade() { Val = 95000 },
            new D.SaturationModulation() { Val = 105000 })
          { Val = D.SchemeColorValues.PhColor }),
          new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
        {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
        },
        new D.Outline(
          new D.SolidFill(
          new D.SchemeColor(
            new D.Shade() { Val = 95000 },
            new D.SaturationModulation() { Val = 105000 })
          { Val = D.SchemeColorValues.PhColor }),
          new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
        {
          Width = 9525,
          CapType = D.LineCapValues.Flat,
          CompoundLineType = D.CompoundLineValues.Single,
          Alignment = D.PenAlignmentValues.Center
        }),
        new D.EffectStyleList(
        new D.EffectStyle(
          new D.EffectList(
          new D.OuterShadow(
            new D.RgbColorModelHex(
            new D.Alpha() { Val = 38000 })
            { Val = "000000" })
          { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
        new D.EffectStyle(
          new D.EffectList(
          new D.OuterShadow(
            new D.RgbColorModelHex(
            new D.Alpha() { Val = 38000 })
            { Val = "000000" })
          { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
        new D.EffectStyle(
          new D.EffectList(
          new D.OuterShadow(
            new D.RgbColorModelHex(
            new D.Alpha() { Val = 38000 })
            { Val = "000000" })
          { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
        new D.BackgroundFillStyleList(
        new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
        new D.GradientFill(
          new D.GradientStopList(
          new D.GradientStop(
            new D.SchemeColor(new D.Tint() { Val = 50000 },
              new D.SaturationModulation() { Val = 300000 })
            { Val = D.SchemeColorValues.PhColor })
          { Position = 0 },
          new D.GradientStop(
            new D.SchemeColor(new D.Tint() { Val = 50000 },
              new D.SaturationModulation() { Val = 300000 })
            { Val = D.SchemeColorValues.PhColor })
          { Position = 0 },
          new D.GradientStop(
            new D.SchemeColor(new D.Tint() { Val = 50000 },
              new D.SaturationModulation() { Val = 300000 })
            { Val = D.SchemeColorValues.PhColor })
          { Position = 0 }),
          new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
        new D.GradientFill(
          new D.GradientStopList(
          new D.GradientStop(
            new D.SchemeColor(new D.Tint() { Val = 50000 },
              new D.SaturationModulation() { Val = 300000 })
            { Val = D.SchemeColorValues.PhColor })
          { Position = 0 },
          new D.GradientStop(
            new D.SchemeColor(new D.Tint() { Val = 50000 },
              new D.SaturationModulation() { Val = 300000 })
            { Val = D.SchemeColorValues.PhColor })
          { Position = 0 }),
          new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
        { Name = "Office" });

      theme1.Append(themeElements1);
      theme1.Append(new D.ObjectDefaults());
      theme1.Append(new D.ExtraColorSchemeList());

      themePart1.Theme = theme1;
      return themePart1;

    }
    
  }
}
