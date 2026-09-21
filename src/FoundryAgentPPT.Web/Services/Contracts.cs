namespace FoundryAgentPPT.Web.Services;

public sealed record CreatePresentationRequest(
    string Topic,
    string OutputPath);

public sealed record CreatePresentationResponse(string OutputPath, int SlideCount);

public sealed class FoundryOptions
{
    public const string SectionName = "Foundry";

    public string ProjectEndpoint { get; init; } = string.Empty;

    public string AgentName { get; init; } = string.Empty;
}

public sealed class SharePointOptions
{
    public const string SectionName = "SharePoint";

    public string DriveId { get; init; } = string.Empty;
}
