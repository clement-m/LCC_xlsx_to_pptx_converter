using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using F = System.IO;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using Drawing = DocumentFormat.OpenXml.Drawing;

using OFFICE2010 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace LCC_xlsx_to_pptx_converter.Class
{
  public static class PPTXClass
  {
    public static void CreatePresentation(string path)
    {
      int position = 2;
      string layoutName = "test name layout 1";

      if (System.IO.File.Exists(path))
      {
        System.IO.File.Delete(path);
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

    public static SlideId getPrevSlideId(SlideIdList slideIdList, int position)
    {
      // Find the highest slide ID in the current list.
      SlideId prevSlideId = null;

      foreach (SlideId slideId in slideIdList.ChildElements)
      {
        position--;
        if (position == 0)
        {
          prevSlideId = slideId;
        }

      }

      return prevSlideId;
    }

    public static uint getMaxSlideId(SlideIdList slideIdList)
    {
      // Find the highest slide ID in the current list.
      uint maxSlideId = 1;

      foreach (SlideId slideId in slideIdList.ChildElements)
      {
        if (slideId.Id > maxSlideId)
        {
          maxSlideId = slideId.Id;
        }
      }

      maxSlideId++;

      return maxSlideId;
    }

    public static void InsertImageInLastSlide(string path, string imagePath)
    {
      using (PresentationDocument presentationDocument = PresentationDocument.Open(path, true))
      {
        SlideIdList SlideIdList = presentationDocument.PresentationPart.Presentation.SlideIdList;

        uint maxId = getMaxSlideId(SlideIdList);

        uint actualId = maxId--;

        // Creates an Picture instance and adds its children. 
        P.Picture picture = new P.Picture();

        string embedId = string.Empty;

        embedId = "rId" + (actualId + 915).ToString();

        P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties(
            new P.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 5" + imagePath },
            new P.NonVisualPictureDrawingProperties(new D.PictureLocks() { NoChangeAspect = true }),
            new ApplicationNonVisualDrawingProperties());

        P.BlipFill blipFill = new P.BlipFill();
        Blip blip = new Blip() { Embed = embedId };

        // Creates an BlipExtensionList instance and adds its children 
        BlipExtensionList blipExtensionList = new BlipExtensionList();
        BlipExtension blipExtension = new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

        OFFICE2010.UseLocalDpi useLocalDpi = new OFFICE2010.UseLocalDpi() { Val = false };
        useLocalDpi.AddNamespaceDeclaration("a14",
            "http://schemas.microsoft.com/office/drawing/2010/main");

        blipExtension.Append(useLocalDpi);
        blipExtensionList.Append(blipExtension);
        blip.Append(blipExtensionList);

        Stretch stretch = new Stretch();
        FillRectangle fillRectangle = new FillRectangle();
        stretch.Append(fillRectangle);

        blipFill.Append(blip);
        blipFill.Append(stretch);

        // Creates an ShapeProperties instance and adds its children. 
        P.ShapeProperties shapeProperties = new P.ShapeProperties();

        D.Transform2D transform2D = new D.Transform2D();
        D.Offset offset = new D.Offset() { X = 457200L, Y = 1524000L };
        D.Extents extents = new D.Extents() { Cx = 8229600L, Cy = 5029200L };

        transform2D.Append(offset);
        transform2D.Append(extents);

        D.PresetGeometry presetGeometry = new D.PresetGeometry() { Preset = D.ShapeTypeValues.Rectangle };
        D.AdjustValueList adjustValueList = new D.AdjustValueList();

        presetGeometry.Append(adjustValueList);

        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        picture.Append(nonVisualPictureProperties);
        picture.Append(blipFill);
        picture.Append(shapeProperties);
        
        //slide.CommonSlideData.ShapeTree.AppendChild(picture);

        // Generates content of imagePart. 
        //ImagePart imagePart = slide.SlidePart.AddNewPart<ImagePart>("png", embedId);
        //F.FileStream fileStream = new F.FileStream(imagePath, F.FileMode.Open);
        //imagePart.FeedData(fileStream);
        //fileStream.Close();
      }
    }

    // Insert a slide into the specified presentation.
    public static void InsertNewSlide(string presentationFile, int position, string slideTitle)
    {
      // Open the source document as read/write. 
      using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
      {
        // Pass the source document and the position and title of the slide to be inserted to the next method.
        InsertNewSlide(presentationDocument, position, slideTitle);
      }
    }
    public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
    {

      if (presentationDocument == null)
      {
        throw new ArgumentNullException("presentationDocument");
      }

      if (slideTitle == null)
      {
        throw new ArgumentNullException("slideTitle");
      }

      PresentationPart presentationPart = presentationDocument.PresentationPart;

      // Verify that the presentation is not empty.
      if (presentationPart == null)
      {
        throw new InvalidOperationException("The presentation document is empty.");
      }

      // Declare and instantiate a new slide.
      Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
      uint drawingObjectId = 1;

      // Construct the slide content.            
      // Specify the non-visual properties of the new slide.
      P.NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new P.NonVisualGroupShapeProperties());
      nonVisualProperties.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 1, Name = "" };
      nonVisualProperties.NonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
      nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

      // Specify the group shape properties of the new slide.
      slide.CommonSlideData.ShapeTree.AppendChild(new P.GroupShapeProperties());

      // Declare and instantiate the title shape of the new slide.
      P.Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());

      drawingObjectId++;

      // Specify the required shape properties for the title shape. 
      titleShape.NonVisualShapeProperties = new P.NonVisualShapeProperties
          (new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
          new P.NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
          new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
      titleShape.ShapeProperties = new P.ShapeProperties();

      // Specify the text of the title shape.
      titleShape.TextBody = new P.TextBody(new Drawing.BodyProperties(),
              new Drawing.ListStyle(),
              new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));

      // Declare and instantiate the body shape of the new slide.
      P.Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
      drawingObjectId++;

      // Specify the required shape properties for the body shape.
      bodyShape.NonVisualShapeProperties = new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
              new P.NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
              new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
      bodyShape.ShapeProperties = new P.ShapeProperties();

      // Specify the text of the body shape.
      bodyShape.TextBody = new P.TextBody(
        new Drawing.BodyProperties(),
              new Drawing.ListStyle(),
              new Drawing.Paragraph());

      // Create the slide part for the new slide.
      SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

      // Save the new slide part.
      slide.Save(slidePart);

      // Modify the slide ID list in the presentation part.
      // The slide ID list should not be null.
      SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

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
    }
    private static void CreatePresentationParts(PresentationPart presentationPart, string masterId = "rId1", string slideId = "rId2")
    {
      SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = masterId });
      SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = slideId });
      SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
      NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
      DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

      presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

      SlidePart slidePart1;
      SlideLayoutPart slideLayoutPart1;
      SlideMasterPart slideMasterPart1;
      ThemePart themePart1;


      slidePart1 = CreateSlidePart(presentationPart, slideId);
      slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
      slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
      themePart1 = CreateTheme(slideMasterPart1);

      slideMasterPart1.AddPart(slideLayoutPart1, masterId);
      presentationPart.AddPart(slideMasterPart1, masterId);
      presentationPart.AddPart(themePart1, "rId5");
    }
    private static SlidePart CreateSlidePart(PresentationPart presentationPart, string slideId = "rId2")
    {
      SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>(slideId);
      slidePart1.Slide = new Slide(
              new CommonSlideData(
                  new ShapeTree(
                      new P.NonVisualGroupShapeProperties(
                          new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                          new P.NonVisualGroupShapeDrawingProperties(),
                          new ApplicationNonVisualDrawingProperties()),
                      new P.GroupShapeProperties(new TransformGroup()),
                      new P.Shape(
                          new P.NonVisualShapeProperties(
                              new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                              new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                              new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                          new P.ShapeProperties(),
                          new P.TextBody(
                              new BodyProperties(),
                              new ListStyle(),
                              new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
              new ColorMapOverride(new MasterColorMapping()));
      return slidePart1;
    }
    private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1, string masterId = "rId1")
    {
      SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>(masterId);
      SlideLayout slideLayout = new SlideLayout(
      new CommonSlideData(new ShapeTree(
        new P.NonVisualGroupShapeProperties(
        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
        new P.NonVisualGroupShapeDrawingProperties(),
        new ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new TransformGroup()),
        new P.Shape(
        new P.NonVisualShapeProperties(
          new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
          new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
          new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
        new P.ShapeProperties(),
        new P.TextBody(
          new BodyProperties(),
          new ListStyle(),
          new Paragraph(new EndParagraphRunProperties()))))),
      new ColorMapOverride(new MasterColorMapping()));
      slideLayoutPart1.SlideLayout = slideLayout;
      return slideLayoutPart1;
    }
    private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1, string masterId = "rId1")
    {
      SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>(masterId);
      SlideMaster slideMaster = new SlideMaster(
      new CommonSlideData(new ShapeTree(
        new P.NonVisualGroupShapeProperties(
        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
        new P.NonVisualGroupShapeDrawingProperties(),
        new ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new TransformGroup()),
        new P.Shape(
        new P.NonVisualShapeProperties(
          new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
          new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
          new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
        new P.ShapeProperties(),
        new P.TextBody(
          new BodyProperties(),
          new ListStyle(),
          new Paragraph())))),
      new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
      new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
      new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
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
