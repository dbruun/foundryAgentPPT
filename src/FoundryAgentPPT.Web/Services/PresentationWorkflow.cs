namespace FoundryAgentPPT.Web.Services;

public sealed class PresentationWorkflow(
    FoundryAgentService agent,
    PowerPointService powerPoint,
    SharePointUploadService sharePoint)
{
    public async Task<CreatePresentationResponse> CreateAsync(
        CreatePresentationRequest request,
        CancellationToken cancellationToken)
    {
        var outline = await agent.CreateOutlineAsync(request.Topic, cancellationToken);
        var presentation = powerPoint.Create(outline);

        await sharePoint.UploadAsync(request.OutputPath, presentation, cancellationToken);
        return new CreatePresentationResponse(request.OutputPath, powerPoint.CountSlides(outline));
    }
}
