using FoundryAgentPPT.Web.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.HttpOverrides;
using System.Net;

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<FoundryOptions>(builder.Configuration.GetSection(FoundryOptions.SectionName));
builder.Services.Configure<SharePointOptions>(builder.Configuration.GetSection(SharePointOptions.SectionName));
var entraId = builder.Configuration.GetSection(EntraIdOptions.SectionName).Get<EntraIdOptions>()
    ?? throw new InvalidOperationException("EntraId configuration is required.");
if (string.IsNullOrWhiteSpace(entraId.Authority) || string.IsNullOrWhiteSpace(entraId.Audience))
{
    throw new InvalidOperationException("EntraId:Authority and EntraId:Audience must be configured.");
}
builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.Authority = entraId.Authority;
        options.Audience = entraId.Audience;
    });
builder.Services.AddAuthorization();
var trustedProxy = builder.Configuration["ForwardedHeaders:TrustedProxy"];
if (!string.IsNullOrWhiteSpace(trustedProxy))
{
    if (!IPAddress.TryParse(trustedProxy, out var trustedProxyAddress))
    {
        throw new InvalidOperationException("ForwardedHeaders:TrustedProxy must be a valid IP address.");
    }

    builder.Services.Configure<ForwardedHeadersOptions>(options =>
    {
        options.ForwardedHeaders = ForwardedHeaders.XForwardedProto;
        options.KnownProxies.Add(trustedProxyAddress);
    });
}
builder.Services.AddSingleton<SharePointUploadService>();
builder.Services.AddSingleton<FoundryAgentService>();
builder.Services.AddSingleton<PowerPointService>();
builder.Services.AddSingleton<PresentationWorkflow>();

var app = builder.Build();

if (!string.IsNullOrWhiteSpace(trustedProxy))
{
    app.UseForwardedHeaders();
}
app.UseHttpsRedirection();
app.UseDefaultFiles();
app.UseStaticFiles();
app.UseAuthentication();
app.UseAuthorization();

app.MapPost("/api/presentations", async (
    CreatePresentationRequest request,
    PresentationWorkflow workflow,
    CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(request.Topic) || string.IsNullOrWhiteSpace(request.OutputPath))
    {
        return Results.ValidationProblem(new Dictionary<string, string[]>
        {
            ["request"] = ["A topic and SharePoint output path are required."]
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
}).RequireAuthorization();

app.MapFallbackToFile("index.html");

app.Run();

public partial class Program;
