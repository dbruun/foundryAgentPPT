# Foundry SharePoint presentation agent

An ASP.NET Core app that asks an existing Foundry IQ-enabled agent to reason over its configured SharePoint knowledge sources through the Microsoft Agent Framework, generates a slide outline, and uploads the resulting `.pptx` to SharePoint with Microsoft Graph.

## Configure and run

Authenticate locally with `az login` (or configure a managed identity in Azure), then set:

```bash
export Foundry__ProjectEndpoint="https://<your-project>.services.ai.azure.com/api/projects/<your-project>"
export Foundry__AgentName="<existing-agent-name>"
export SharePoint__DriveId="<document-library-drive-id>"
dotnet run --project src/FoundryAgentPPT.Web
```

Open the displayed URL and provide a topic and destination `.pptx` path. The Foundry agent must already be configured with a Foundry IQ connection to the relevant SharePoint knowledge sources. The application identity needs Microsoft Graph permission to write to the configured document library plus access to the Foundry project and agent.