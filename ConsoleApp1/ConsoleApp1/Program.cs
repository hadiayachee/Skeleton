using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Paragraph = DocumentFormat.OpenXml.Drawing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Drawing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using RunProperties = DocumentFormat.OpenXml.Drawing.RunProperties;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using TextBody = DocumentFormat.OpenXml.Presentation.TextBody;


class Program
{
    static void Main(string[] args)
    {
        string filePath = "C:/Users/Hadi/Desktop/AUXI TASK/ConsoleApp1/auxi C# Interview.pptx";

        try
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                // Get the presentation part
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                // Get the input and output slide parts
                //RiD3 is the id of slide in pptx to know which slide that take effects
                SlidePart inputSlidePart = presentationPart.GetPartById("rId3") as SlidePart;
               // SlidePart outputSlidePart = presentationPart.GetPartById("rId3") as SlidePart;

                // Ensure the input slide part exists before modifying
                if (inputSlidePart != null)
                {
                    // Change title of input slide and center text
                    ChangeSlideTitle(inputSlidePart, "Output Slide");


                    // Remove TextBox  from the slide
                    RemoveTextBox(inputSlidePart, "TextBox 2");
                    RemoveTextBox(inputSlidePart, "TextBox 3");
                    RemoveTextBox(inputSlidePart, "TextBox 4");
                    RemoveTextBox(inputSlidePart, "TextBox 5");
                    //change the size of shapes
                    SetShapeSize(inputSlidePart, "Arrow: Pentagon 7", widthInches: 3.04, heightInches: 1.58);
                    SetShapeSize(inputSlidePart, "Arrow: Chevron 8", widthInches: 3.04, heightInches: 1.58);
                    SetShapeSize(inputSlidePart, "Arrow: Chevron 9", widthInches: 3.04, heightInches: 1.58);
                    SetShapeSize(inputSlidePart, "Arrow: Chevron 11", widthInches: 3.04, heightInches: 1.58);

                    

                   
                    //male them in same level with specfic space the value is in negative because its shifted to left
                    PositionShapesOnSameLevel(inputSlidePart,
                    startingX: 0.5, //  coordinate in inches
                    startingY: 1.7, // coordinate in inches
                    horizontalSpacing: -0.7, // spacing between shapes in inches
                   "Arrow: Pentagon 7", "Arrow: Chevron 8", "Arrow: Chevron 9", "Arrow: Chevron 11");

                    PositionShapesOnSameLevel(inputSlidePart,
                startingX: 0.5, //  coordinate in inches
                startingY: 4.0, //  coordinate in inches
                horizontalSpacing: 1.0, //  spacing between shapes in inches
               "TextBox 26", "TextBox 27", "TextBox 28", "TextBox 15");

                    
                    AddTextToPentagon(inputSlidePart, "Arrow: Pentagon 7", "Begin");
                    AddTextToChevron(inputSlidePart, "Arrow: Chevron 8", "Step 1");
                    AddTextToChevron(inputSlidePart, "Arrow: Chevron 9", "Step 2");
                    AddTextToChevron(inputSlidePart, "Arrow: Chevron 11", "Step 3");


                    //intialize the parameters of textbox as named from pptx
                    Shape textBox15 = inputSlidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(shape =>
                       shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == "TextBox 15");
                    Shape textBox26 = inputSlidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(shape =>
                      shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == "TextBox 26");
                    Shape textBox27 = inputSlidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(shape =>
                      shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == "TextBox 27");
                    Shape textBox28 = inputSlidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(shape =>
                      shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == "TextBox 28");

                    //here we can put in if statment many textbox but i put only textbox15!null
                    if (textBox15 != null)
                    {
                        RemoveBoldAndUnderline(textBox15);
                        ChangeTextBoxFont(textBox15, "Beirut");
                        ChangeTextBoxFont(textBox26, "Beirut");
                        ChangeTextBoxFont(textBox27, "Beirut");
                        ChangeTextBoxFont(textBox28, "Beirut");
                        ChangeTextBoxBulletStyle(textBox28, "ListBullet");
                        ChangeTextBoxBulletStyle2(textBox28, "•");

                        ChangeTextBoxBulletStyle(textBox26, "ListBullet");
                        ChangeTextBoxBulletStyle2(textBox26, "•");

                        ChangeTextBoxBulletStyle(textBox27, "ListBullet");
                        ChangeTextBoxBulletStyle2(textBox27, "•");
                        ChangeTextBoxBulletStyle2(textBox15, "•");

                    }



                }

                // Save changes
                presentationPart.Presentation.Save();
                Console.WriteLine("Changes saved successfully.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    //This method to change the title value.txt
    static void ChangeSlideTitle(SlidePart slidePart, string newTitle)
    {
        Slide slide = slidePart.Slide;


        // Find the title shape in the slide named from pptx
        Shape titleShape = slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(s => s.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value == "Title 1");

        if (titleShape != null)
        {
            TextBody textBody = titleShape.TextBody;
            if (textBody != null)
            {
                // Find the first text paragraph
                Paragraph paragraph = textBody.Elements<Paragraph>().FirstOrDefault();
                if (paragraph != null)
                {
                    // Find the first text run in the paragraph
                    Run run = paragraph.Elements<Run>().FirstOrDefault();
                    if (run != null)
                    {
                        // Change text and font properties
                        Text text = run.Text;
                        text.Text = newTitle;

                        run.RunProperties = new RunProperties(new A.LatinFont() { Typeface = "Beirut" });

                        // Set alignment to top-center but is dosnt work because there is an error in packges versions on my laptop its the same for other in chevon
                      /*  paragraph.ParagraphProperties = new ParagraphProperties(
                            new A.Alignment() { Vertical = A.TextVerticalValues.Top, Horizontal = A.TextAlignmentTypeValues.Center }
                        );*/
                    }
                }
            }
        }
    }

    //This method to delete the textboxes
    static void RemoveTextBox(SlidePart slidePart, string textBoxName)
    {
        Shape textBox = slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(shape =>
            shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == textBoxName);

        if (textBox != null)
        {
            // Remove the textbox shape
            textBox.Remove();
        }
    
}
    //This method to add text value to cheveron
    static void AddTextToChevron(SlidePart slidePart, string chevronName, string textToAdd)
    {
        Shape chevron = slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(shape =>
            shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == chevronName);

        if (chevron != null)
        {
            // Create a new TextBody element
            TextBody textBody = new TextBody();

            // Create a new Paragraph and Run with the specified text
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Text text = new Text();
            text.Text = textToAdd;
            run.Append(text);
            paragraph.Append(run);

            // Append the Paragraph to the TextBody
            textBody.Append(paragraph);

            // Set the TextBody for the Chevron shape
            chevron.TextBody = textBody;
        }
    }

    //same as chevron but for Pentagon
    static void AddTextToPentagon(SlidePart slidePart, string pentagonName, string textToAdd)
    {
        Shape pentagon = slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(shape =>
            shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == pentagonName);

        if (pentagon != null)
        {
            // Create a new TextBody element
            TextBody textBody = new TextBody();

            // Create a new Paragraph and Run with the specified text
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Text text = new Text();
            text.Text = textToAdd;
            run.Append(text);
            paragraph.Append(run);

            // Append the Paragraph to the TextBody
            textBody.Append(paragraph);

            // Set the TextBody for the Pentagon shape
            pentagon.TextBody = textBody;
        }
    }
    //This function for make the elements on same level behind each other
    static void PositionShapesOnSameLevel(SlidePart slidePart,
    double startingX, double startingY, double horizontalSpacing,
    params string[] shapeNames)
    {
        double currentX = startingX;
        double yCoordinate = startingY;

        foreach (string shapeName in shapeNames)
        {
            Shape shape = slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(s =>
                s.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == shapeName);

            if (shape != null)
            {
                SetShapePosition(shape, currentX, yCoordinate);

                currentX += shape.ShapeProperties.Transform2D.Extents.Cx.Value / 914400.0 + horizontalSpacing;
            }
        }
    }

    static void SetShapePosition(Shape shape, double xInches, double yInches)
    {
        long xEmu = (long)(xInches * 914400); // Convert inches to EMUs
        long yEmu = (long)(yInches * 914400); // Convert inches to EMUs

        shape.ShapeProperties.Transform2D.Offset.X = xEmu;
        shape.ShapeProperties.Transform2D.Offset.Y = yEmu;
    }
    static void SetShapeSize(SlidePart slidePart, string shapeName, double widthInches, double heightInches)
    {
        Shape shape = slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(s =>
            s.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value == shapeName);

        if (shape != null)
        {
            long widthEmu = (long)(widthInches * 914400); // Convert inches to EMUs
            long heightEmu = (long)(heightInches * 914400); // Convert inches to EMUs

            // Set the size of the shape
            shape.ShapeProperties.Transform2D.Extents.Cx = widthEmu;
            shape.ShapeProperties.Transform2D.Extents.Cy = heightEmu;
        }
    }

    //This function to remove bold and unerlines
    static void RemoveBoldAndUnderline(Shape textBox)
    {
        TextBody textBody = textBox.TextBody;
        if (textBody != null)
        {
            foreach (var paragraph in textBody.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        // Clear bold and underline formats
                        run.RunProperties = new RunProperties();
                    }
                }
            }
        }
    }
    
   
    //This funtion to change the type of font 
    static void ChangeTextBoxFont(Shape textBox, string fontName)
    {
        TextBody textBody = textBox.TextBody;
        if (textBody != null)
        {
            foreach (var paragraph in textBody.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        // Set the desired font for the text
                        run.RunProperties = new RunProperties(new A.LatinFont() { Typeface = fontName });
                    }
                }
            }
        }
    }
    //This function to remove all bullets
    static void ChangeTextBoxBulletStyle(Shape textBox, string bulletStyle)
    {
        TextBody textBody = textBox.TextBody;
        if (textBody != null)
        {
            foreach (var paragraph in textBody.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        // Apply the desired bullet style to the paragraph
                        paragraph.ParagraphProperties = new ParagraphProperties(
                            new ParagraphStyleId() { Val = bulletStyle }
                        );
                    }
                }
            }
        }
    }
    //This function to add all bullets as dot and the Flag to add them for 1 time 
    static void ChangeTextBoxBulletStyle2(Shape textBox, string bulletSymbol)
    {
        TextBody textBody = textBox.TextBody;
        if (textBody != null)
        {
            bool hasBulletPoints = false; // Flag to track existing bullet points

            foreach (var paragraph in textBody.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        // Check if the text contains the specified bullet symbol
                        if (text.Text.Contains(bulletSymbol))
                        {
                            hasBulletPoints = true;
                            break; // No need to continue checking
                        }
                    }
                }

                if (!hasBulletPoints)
                {
                    // Create a Run element for the bullet symbol
                    var bulletRun = new Run(new RunProperties(new RunFonts() { Ascii = "Symbol" }), new Text(bulletSymbol));

                    // Create a Run element for the space
                    var spaceRun = new Run(new Text("   "));

                    // Insert the bullet run and space run at the beginning of the paragraph
                    paragraph.InsertAt(bulletRun, 0);
                    paragraph.InsertAt(spaceRun, 1);
                }

                // Reset the flag for the next paragraph
                hasBulletPoints = false;
            }
        }
    }


}
