const DEFAULT_AI_ENDPOINT = 'https://cog-graphshell-ai-4a27a.openai.azure.com/';
const DEFAULT_AI_DEPLOYMENT = 'gpt-4-mini';
const DEFAULT_AI_KEY = '';

// Helper to load Azure OpenAI credentials
function getAzureCredentials() {
  return {
    endpoint: localStorage.getItem('graphshell_ai_endpoint') || DEFAULT_AI_ENDPOINT,
    deployment: localStorage.getItem('graphshell_ai_deployment') || DEFAULT_AI_DEPLOYMENT,
    apiKey: localStorage.getItem('graphshell_ai_key') || DEFAULT_AI_KEY
  };
}

// System prompt to guide the LLM model to translate natural language to GraphShell commands
const SYSTEM_PROMPT = `You are the Azure AI Copilot for GraphShell, a terminal command interface for Microsoft Graph. Your goal is to translate natural language user queries into GraphShell CLI commands.

Available commands are:
1. help - Show commands listing
2. me - View current user profile
3. users [--list] [--create] [--delete] [--department <name>]
   - --create requires: -n "Name" -j "Job" -d "Dept" (or --name, --jobTitle, --department)
   - --delete requires: -id "user-id-or-mail"
4. groups - List AD Directory groups
5. mail [--list] [--send]
   - --send requires: --to "email" --subject "sub" --body "text"
6. calendar [--list] [--add]
   - --add requires: --subject "sub" --start "YYYY-MM-DDTHH:MM" --end "YYYY-MM-DDTHH:MM"
7. drive [--ls] [--mkdir] [--upload]
   - --mkdir requires: --name "name"
   - --upload requires: --name "name"
8. request [GET/POST/DELETE] [endpoint]
   - Execute raw API requests. E.g., "request GET /me/messages"
9. clear - Clear screen

Reply ONLY with a valid JSON containing:
1. "explanation": A concise, friendly summary of what action you took.
2. "command": The exact GraphShell command string to execute. If no command matches, set this to null.

Format example:
{
  "explanation": "Retrieving your Outlook inbox messages.",
  "command": "mail --list"
}`;

// Main Prompt Handler (Calls Azure OpenAI)
export async function runAiPrompt(promptText, terminal = null) {
  const { endpoint, deployment, apiKey } = getAzureCredentials();
  
  if (!endpoint || !deployment || !apiKey) {
    throw new Error("Azure OpenAI credentials not configured. Open settings to save configurations.");
  }
  
  const cleanedEndpoint = endpoint.replace(/\/$/, '');
  const url = `${cleanedEndpoint}/openai/deployments/${deployment}/chat/completions?api-version=2024-02-15-preview`;
  
  const requestBody = {
    messages: [
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: promptText }
    ],
    temperature: 0.1,
    max_tokens: 400
  };
  
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'api-key': apiKey,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(requestBody)
  });
  
  if (!res.ok) {
    const errText = await res.text().catch(() => '');
    throw new Error(`Azure AI HTTP Error ${res.status}: ${res.statusText}. Details: ${errText}`);
  }
  
  const data = await res.json();
  const messageContent = data.choices?.[0]?.message?.content;
  if (!messageContent) {
    throw new Error("Empty response received from Azure OpenAI.");
  }
  
  // Parse the JSON output from assistant
  try {
    return JSON.parse(messageContent.trim());
  } catch (parseErr) {
    console.warn("Raw assistant reply was not valid JSON:", messageContent);
    return {
      explanation: messageContent,
      command: null
    };
  }
}

const ERROR_EXPLAINER_PROMPT = `You are the Azure AI Copilot for GraphShell, a technical debugger for Microsoft Graph API requests.
The user's Microsoft Graph REST query has failed.
Analyze the provided API endpoint, the user's account type (e.g. Personal MSA, Work/School tenant), and the Graph error message.
Explain clearly:
1. Why this API is not supported or why the query failed (e.g., Personal MSA accounts do not support directory search /organization or directory write commands /users).
2. The exact workaround or solution (e.g. switch to a Work/School tenant account, consent to Organization.Read.All, etc.).

Provide a concise, friendly, and structured Markdown output (using bullet points or clear sections) that explains the cause and the solution. Do not include JSON formatting or HTML tags, just pure Markdown text.`;

export async function explainApiError(endpoint, errorMsg, accountType) {
  const { endpoint: aiEndpoint, deployment, apiKey } = getAzureCredentials();
  
  if (!aiEndpoint || !deployment || !apiKey) {
    return "Error Analysis: Azure OpenAI credentials not configured. Open settings to save configurations.";
  }
  
  const cleanedEndpoint = aiEndpoint.replace(/\/$/, '');
  const url = `${cleanedEndpoint}/openai/deployments/${deployment}/chat/completions?api-version=2024-02-15-preview`;
  
  const promptText = `API Endpoint called: ${endpoint}
Error Message received: ${errorMsg}
User Account Type: ${accountType}`;
  
  const requestBody = {
    messages: [
      { role: "system", content: ERROR_EXPLAINER_PROMPT },
      { role: "user", content: promptText }
    ],
    temperature: 0.3,
    max_tokens: 600
  };
  
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'api-key': apiKey,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(requestBody)
    });
    
    if (!res.ok) return `Could not fetch AI analysis. (HTTP ${res.status} ${res.statusText})`;
    
    const data = await res.json();
    return data.choices?.[0]?.message?.content || "No analysis generated.";
  } catch (err) {
    return `Could not connect to Azure OpenAI for error analysis: ${err.message}`;
  }
}
