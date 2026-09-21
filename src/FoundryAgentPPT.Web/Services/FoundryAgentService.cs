using Azure.AI.Projects;
using Azure.Identity;
using Microsoft.Agents.AI;
using Microsoft.Extensions.Options;

namespace FoundryAgentPPT.Web.Services;

public sealed class FoundryAgentService(IOptions<FoundryOptions> options)
{
    private readonly FoundryOptions options = options.Value;

    public async Task<string> CreateOutlineAsync(
        string topic,
        IReadOnlyList<SourceDocument> documents,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(options.ProjectEndpoint) || string.IsNullOrWhiteSpace(options.AgentName))
        {
            throw new InvalidOperationException("Foundry:ProjectEndpoint and Foundry:AgentName must be configured.");
        }

        var project = new AIProjectClient(new Uri(options.ProjectEndpoint), new DefaultAzureCredential());
        var agentRecord = await project.Agents.GetAgentAsync(options.AgentName, cancellationToken);
        var agent = project.AsAIAgent(agentRecord.Value);
        var sources = string.Join(
            "\n\n",
            documents.Select(document => $"SOURCE: {document.Path}\n{document.Text}"));
        var response = await agent.RunAsync(
            $"Create a concise slide outline for '{topic}' using only the sources below. " +
            "Return one slide per line using the format: Title | bullet one; bullet two; bullet three.\n\n" +
            sources,
            cancellationToken: cancellationToken);

        return response.Text;
    }
}
