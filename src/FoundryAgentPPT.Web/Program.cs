using FoundryAgentPPT.Web.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<FoundryOptions>(builder.Configuration.GetSection(FoundryOptions.SectionName));
builder.Services.Configure<SharePointOptions>(builder.Configuration.GetSection(SharePointOptions.SectionName));
builder.Services.AddSingleton<SharePointDocumentService>();
builder.Services.AddSingleton<FoundryAgentService>();
builder.Services.AddSingleton<PowerPointService>();
builder.Services.AddSingleton<PresentationWorkflow>();

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapPost("/api/presentations", async (
    CreatePresentationRequest request,
    PresentationWorkflow workflow,
    CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(request.Topic) || request.DocumentPaths.Count == 0)
    {
        return Results.ValidationProblem(new Dictionary<string, string[]>
        {
            ["request"] = ["A topic and at least one SharePoint document path are required."]
        });
    }

    try
    {
        var result = await workflow.CreateAsync(request, cancellationToken);
        return Results.Ok(result);
    }
    catch (InvalidOperationException exception)
    {
        return Results.Problem(exception.Message, statusCode: StatusCodes.Status400BadRequest);
    }
});

app.MapFallbackToFile("index.html");

app.Run();

public partial class Program;
