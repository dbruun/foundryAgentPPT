using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FoundryAgentPPT.Web.Services;

public sealed class PowerPointService
{
    private const long TitleX = 457200;
    private const long TitleY = 274638;
    private const long TitleWidth = 11239500;
    private const long TitleHeight = 914400;
    private const int TitleFontSize = 3200;
    private const long ContentY = 1371600;
    private const long ContentHeight = 4572000;
    private const int ContentFontSize = 1800;

    public MemoryStream Create(string outline)
    {
        var stream = new MemoryStream();

        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, true))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation(
                new P.SlideMasterIdList(),
                new P.SlideIdList(),
                new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 },
                new P.NotesSize { Cx = 6858000, Cy = 9144000 });

            var layoutPart = CreateLayout(presentationPart);
            var slideIdList = presentationPart.Presentation.GetFirstChild<P.SlideIdList>()!;
            uint id = 256;
            foreach (var slide in Parse(outline))
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new P.Slide(
                    new P.CommonSlideData(CreateShapeTree(slide)),
                    new P.ColorMapOverride(new A.MasterColorMapping()));
                slidePart.AddPart(layoutPart);

                slideIdList.Append(new P.SlideId
                {
                    Id = id++,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                });
            }

            presentationPart.Presentation.Save();
        }

        stream.Position = 0;
        return stream;
    }

    public int CountSlides(string outline) => Parse(outline).Count;

    private static List<SlideContent> Parse(string outline) =>
        outline
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(line => line.Split('|', 2, StringSplitOptions.TrimEntries))
            .Select(parts => new SlideContent(
                parts[0],
                parts.Length == 2
                    ? parts[1].Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    : []))
            .ToList() is { Count: > 0 } slides
                ? slides
                : throw new InvalidOperationException("The agent returned an empty slide outline.");

    private static P.ShapeTree CreateShapeTree(SlideContent content)
    {
        var tree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));

        tree.Append(CreateTextShape(2U, "Title", content.Title, TitleX, TitleY, TitleWidth, TitleHeight, TitleFontSize));
        tree.Append(CreateBulletShape(content));

        return tree;
    }

    private static P.Shape CreateBulletShape(SlideContent content)
    {
        var textBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var bullet in content.Bullets.DefaultIfEmpty(string.Empty))
        {
            textBody.Append(new A.Paragraph(
                new A.Run(
                    new A.RunProperties { FontSize = ContentFontSize },
                    new A.Text($"• {bullet}"))));
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 3U, Name = "Content" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = TitleX, Y = ContentY },
                    new A.Extents { Cx = TitleWidth, Cy = ContentHeight })),
            textBody);
    }

    private static SlideLayoutPart CreateLayout(PresentationPart presentationPart)
    {
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        layoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(CreateShapeTree(new SlideContent(string.Empty, []))),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        masterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(CreateShapeTree(new SlideContent(string.Empty, []))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new P.SlideLayoutIdList(new P.SlideLayoutId
            {
                Id = 2147483648U,
                RelationshipId = masterPart.GetIdOfPart(layoutPart)
            }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));

        var masterIdList = presentationPart.Presentation!.SlideMasterIdList
            ?? throw new InvalidOperationException("The presentation is missing its slide-master list.");
        masterIdList.Append(new P.SlideMasterId
        {
            Id = 2147483648U,
            RelationshipId = presentationPart.GetIdOfPart(masterPart)
        });

        return layoutPart;
    }

    private static P.Shape CreateTextShape(
        uint id,
        string name,
        string text,
        long x,
        long y,
        long cx,
        long cy,
        int fontSize) =>
        new(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy })),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.RunProperties { FontSize = fontSize },
                        new A.Text(text)))));

    private sealed record SlideContent(string Title, IReadOnlyList<string> Bullets);
}
