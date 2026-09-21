namespace FoundryAgentPPT.Web.Services;

public sealed class PresentationWorkflow(
    SharePointDocumentService sharePoint,
    FoundryAgentService agent,
    PowerPointService powerPoint)
{
    public async Task<CreatePresentationResponse> CreateAsync(
        CreatePresentationRequest request,
        CancellationToken cancellationToken)
    {
        var documents = await sharePoint.ReadAsync(request.DocumentPaths, cancellationToken);
        var outline = await agent.CreateOutlineAsync(request.Topic, documents, cancellationToken);
        var presentation = powerPoint.Create(outline);

        await sharePoint.UploadAsync(request.OutputPath, presentation, cancellationToken);
        return new CreatePresentationResponse(request.OutputPath, powerPoint.CountSlides(outline));
    }
}
