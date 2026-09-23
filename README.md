# Foundry SharePoint presentation agent

An ASP.NET Core application that asks an existing Microsoft Foundry agent to research its configured SharePoint knowledge sources through Foundry IQ, turns the resulting slide outline into a PowerPoint presentation, and uploads the `.pptx` to a SharePoint document library.

The application deliberately does **not** read source documents from SharePoint. The Foundry agent owns document retrieval through its existing Foundry IQ connection. This keeps the application's SharePoint access limited to writing the finished presentation.

## How it works

1. A signed-in user submits a topic and output path from the browser UI.
2. The API validates the user's Microsoft Entra ID bearer token.
3. The API uses the Microsoft Agent Framework to invoke the existing Foundry agent with a request to research the topic through its configured Foundry IQ knowledge sources.
4. The agent returns a line-based slide outline.
5. The application creates a schema-valid `.pptx` from that outline and uploads it to the configured SharePoint document-library drive through Microsoft Graph.

## Prerequisites

- .NET 10 SDK.
- An existing Foundry agent that has access to the intended SharePoint content through Foundry IQ.
- A Microsoft Entra ID application registration that issues access tokens for this API.
- An identity available to the web application through `DefaultAzureCredential`. Locally, use Azure CLI authentication; in Azure, prefer managed identity or workload identity.
- Microsoft Graph write permission for the application identity on the destination document library.
- Foundry project and agent access for the application identity.

## Configuration

Authenticate locally with `az login` (or configure a managed identity in Azure), then provide the following environment variables:

```bash
export Foundry__ProjectEndpoint="https://<your-project>.services.ai.azure.com/api/projects/<your-project>"
export Foundry__AgentName="<existing-agent-name>"
export SharePoint__DriveId="<document-library-drive-id>"
export EntraId__Authority="https://login.microsoftonline.com/<tenant-id>/v2.0"
export EntraId__Audience="api://<application-client-id>"
dotnet run --project src/FoundryAgentPPT.Web
```

The application redirects HTTP requests to HTTPS before it accepts caller credentials. For local development, trust the .NET development certificate with `dotnet dev-certs https --trust`. In production, terminate TLS at the application or a trusted proxy. When using a proxy, set `ForwardedHeaders__TrustedProxy` to its IP address; the application then trusts only that proxy's forwarded HTTPS scheme.

| Setting | Purpose |
| --- | --- |
| `Foundry__ProjectEndpoint` | Endpoint of the Foundry project containing the existing agent. |
| `Foundry__AgentName` | Name of the existing Foundry IQ-enabled agent. |
| `SharePoint__DriveId` | Microsoft Graph drive ID of the destination SharePoint document library. |
| `EntraId__Authority` | Microsoft Entra ID authority that validates caller tokens. |
| `EntraId__Audience` | Application ID URI or audience expected in caller tokens. |
| `ForwardedHeaders__TrustedProxy` | Optional IP address of the TLS-terminating proxy that forwards the original HTTPS scheme. |

The same values can be supplied through `src/FoundryAgentPPT.Web/appsettings.json` for deployment configuration. Do not commit credentials or access tokens.

## Use the browser UI

Open the URL printed by `dotnet run` and provide:

- A presentation topic.
- The SharePoint-relative output path, such as `Presentations/qbr.pptx`.
- A Microsoft Entra ID access token for the configured API audience.

The browser sends the token only in the `Authorization` header of the presentation request; it does not store the token. The output destination must be within the configured document-library drive.

## API

`POST /api/presentations` requires an `Authorization` header carrying an access token issued for the configured audience:

```http
POST /api/presentations
Content-Type: application/json

{
  "topic": "Quarterly business review",
  "outputPath": "Presentations/qbr.pptx"
}
```

On success, the API returns the SharePoint output path and number of slides:

```json
{
  "outputPath": "Presentations/qbr.pptx",
  "slideCount": 6
}
```

## Security and permissions

The API is protected because a Foundry IQ agent may have access to broader knowledge sources than the destination library. Restrict Entra ID access to users who are permitted to request research from the configured Foundry agent. Use a dedicated Foundry agent and IQ connection when the application must be limited to a particular body of SharePoint content.

The server-side application identity, not the browser token, accesses Foundry and uploads the PowerPoint. Grant it only the least-privileged Graph and Foundry roles necessary. Keep Foundry IQ source permissions managed within Foundry and SharePoint.

## Troubleshooting

- **401 Unauthorized:** Ensure the access token was issued by `EntraId__Authority` and has the expected `EntraId__Audience`.
- **400 configuration error:** Verify all Foundry, SharePoint drive, and Entra ID settings are populated.
- **Foundry request fails:** Confirm the application identity can access the configured project and agent, and that the agent's Foundry IQ connection is healthy.
- **Upload fails:** Confirm `SharePoint__DriveId`, the requested output path, and Microsoft Graph write permissions for the application identity.