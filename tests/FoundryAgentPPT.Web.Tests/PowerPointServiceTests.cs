using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FoundryAgentPPT.Web.Services;

namespace FoundryAgentPPT.Web.Tests;

public sealed class PowerPointServiceTests
{
    [Fact]
    public void Create_creates_a_slide_for_each_outline_line()
    {
        var service = new PowerPointService();

        using var stream = service.Create("""
            Introduction | Purpose; Audience
            Results | Revenue; Growth
            """);
        using var presentation = PresentationDocument.Open(stream, false);

        var slideIdList = presentation.PresentationPart!.Presentation!.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.Equal(2, slideIdList!.Count());
        var validationErrors = new OpenXmlValidator().Validate(presentation).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors.Select(error => error.Description)));
        Assert.Equal(2, service.CountSlides("Introduction | Purpose\nResults | Revenue"));
    }
}
