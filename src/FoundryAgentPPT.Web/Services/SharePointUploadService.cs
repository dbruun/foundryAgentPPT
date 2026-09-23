using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Extensions.Options;

namespace FoundryAgentPPT.Web.Services;

public sealed class SharePointUploadService(IOptions<SharePointOptions> options)
{
    private readonly SharePointOptions options = options.Value;

    public async Task UploadAsync(string path, Stream content, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(options.DriveId))
        {
            throw new InvalidOperationException("A SharePoint document-library drive ID is required.");
        }

        _ = await new GraphServiceClient(
                new DefaultAzureCredential(),
                ["https://graph.microsoft.com/.default"])
            .Drives[options.DriveId]
            .Root
            .ItemWithPath(path)
            .Content
            .PutAsync(content, cancellationToken: cancellationToken)
            ?? throw new InvalidOperationException($"Unable to upload '{path}' to SharePoint.");
    }
}
