using Azure.AI.Projects;
using Azure.Identity;
using Microsoft.Agents.AI;
using Microsoft.Extensions.Options;

namespace FoundryAgentPPT.Web.Services;

public sealed class FoundryAgentService(IOptions<FoundryOptions> options)
{
    private readonly FoundryOptions options = options.Value;
    private readonly Lazy<AIProjectClient> project = new(() => new AIProjectClient(
        new Uri(options.Value.ProjectEndpoint),
        new DefaultAzureCredential()));

    public async Task<string> CreateOutlineAsync(
        string topic,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(options.ProjectEndpoint) || string.IsNullOrWhiteSpace(options.AgentName))
        {
            throw new InvalidOperationException("Foundry:ProjectEndpoint and Foundry:AgentName must be configured.");
        }

        var agentRecord = await project.Value.Agents.GetAgentAsync(options.AgentName, cancellationToken);
        var agent = project.Value.AsAIAgent(agentRecord.Value);
        var response = await agent.RunAsync(
            $"Use your configured Foundry IQ knowledge sources to research '{topic}'. " +
            "Create a concise slide outline. Return one slide per line using the format: " +
            "Title | bullet one; bullet two; bullet three.",
            cancellationToken: cancellationToken);

        return response.Text;
    }
}
