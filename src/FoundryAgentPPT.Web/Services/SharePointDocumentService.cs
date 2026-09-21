using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Options;

namespace FoundryAgentPPT.Web.Services;

public sealed class SharePointDocumentService(IOptions<SharePointOptions> options)
{
    private readonly SharePointOptions options = options.Value;

    public async Task<IReadOnlyList<SourceDocument>> ReadAsync(
        IEnumerable<string> paths,
        CancellationToken cancellationToken)
    {
        var client = CreateClient();
        var documents = new List<SourceDocument>();

        foreach (var path in paths)
        {
            await using var content = await client
                .Drives[options.DriveId]
                .Root
                .ItemWithPath(path)
                .Content
                .GetAsync(cancellationToken: cancellationToken)
                ?? throw new InvalidOperationException($"SharePoint document '{path}' was not found.");

            using var reader = new StreamReader(content);
            documents.Add(new SourceDocument(path, await reader.ReadToEndAsync(cancellationToken)));
        }

        return documents;
    }

    public async Task UploadAsync(string path, Stream content, CancellationToken cancellationToken)
    {
        _ = await CreateClient()
            .Drives[options.DriveId]
            .Root
            .ItemWithPath(path)
            .Content
            .PutAsync(content, cancellationToken: cancellationToken)
            ?? throw new InvalidOperationException($"Unable to upload '{path}' to SharePoint.");
    }

    private GraphServiceClient CreateClient()
    {
        if (string.IsNullOrWhiteSpace(options.DriveId))
        {
            throw new InvalidOperationException("SharePoint:DriveId must be configured.");
        }

        return new GraphServiceClient(new DefaultAzureCredential(), ["https://graph.microsoft.com/.default"]);
    }
}

public sealed record SourceDocument(string Path, string Text);
