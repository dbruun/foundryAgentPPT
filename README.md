# Foundry SharePoint presentation agent

An ASP.NET Core app that retrieves SharePoint documents with Microsoft Graph, asks an existing Microsoft Foundry agent to create a slide outline through the Microsoft Agent Framework, and uploads the generated `.pptx` back to SharePoint.

## Configure and run

Authenticate locally with `az login` (or configure a managed identity in Azure), then set:

```bash
export Foundry__ProjectEndpoint="https://<your-project>.services.ai.azure.com/api/projects/<your-project>"
export Foundry__AgentName="<existing-agent-name>"
export SharePoint__DriveId="<document-library-drive-id>"
dotnet run --project src/FoundryAgentPPT.Web
```

Open the displayed URL and provide a topic, one or more text-document paths in the selected SharePoint document library, and the destination `.pptx` path. The identity needs Microsoft Graph permissions to read and write that library plus access to the Foundry project and agent.