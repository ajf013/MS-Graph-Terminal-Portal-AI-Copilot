import { runAiPrompt } from './aiService.js';

// Visual and JSON update callbacks (populated by app.js)
export let updateVisualPanel = () => {};
export let updateJsonPanel = () => {};
export let getMsalInstance = () => null;
export let logApiTransaction = () => {};

export function registerCallbacks(visualCb, jsonCb, msalGetter, apiLogCb) {
  updateVisualPanel = visualCb;
  updateJsonPanel = jsonCb;
  getMsalInstance = msalGetter;
  logApiTransaction = apiLogCb;
}

// Scope mapping helper to display required permissions for endpoints
function getScopesForEndpoint(endpoint) {
  const clean = endpoint.split('?')[0];
  if (clean === '/me') return ["User.Read"];
  if (clean.startsWith('/users')) return ["User.ReadBasic.All", "User.Read.All"];
  if (clean.startsWith('/groups')) return ["Group.Read.All", "Directory.Read.All"];
  if (clean.includes('/messages') || clean.includes('/sendMail')) return ["Mail.ReadWrite", "Mail.Send"];
  if (clean.includes('/events')) return ["Calendars.ReadWrite"];
  if (clean.includes('/drive')) return ["Files.ReadWrite"];
  return ["User.Read"];
}

// Print a formatted REST request block directly to the terminal history
function logRequestToTerminal(method, url, scopes, terminal) {
  if (!terminal) return;
  terminal.writeLine("┌────────────────────────────────────────────────────────────────────────────┐", "info");
  terminal.writeLine(`│  [GRAPH REST API CALL]`, "info");
  terminal.writeLine(`│  Method:    ${method.padEnd(65)}│`, "info");
  terminal.writeLine(`│  Endpoint:  ${url.padEnd(63)}│`, "info");
  terminal.writeLine(`│  Scopes:    ${scopes.join(', ').padEnd(65)}│`, "info");
  terminal.writeLine("└────────────────────────────────────────────────────────────────────────────┘", "info");
}

// Helper to check authentication and run query on Microsoft Graph REST API
async function executeGraphRequest(endpoint, method = 'GET', payload = null, terminal = null) {
  const msal = getMsalInstance();
  if (!msal) {
    throw new Error("MSAL is not initialized. Make sure Client ID is set.");
  }
  
  // Attempt to acquire token silently
  const accounts = msal.getAllAccounts();
  if (accounts.length === 0) {
    throw new Error("Not logged in. Click 'Connect Account' or type 'login' to sign in.");
  }
  
  const scopes = getScopesForEndpoint(endpoint);
  
  // Log transaction
  if (terminal) {
    logRequestToTerminal(method, `https://graph.microsoft.com/v1.0${endpoint}`, scopes, terminal);
    logApiTransaction(method, `https://graph.microsoft.com/v1.0${endpoint}`, scopes.join(', '), 'PENDING');
  }

  const request = {
    scopes: scopes,
    account: accounts[0]
  };
  
  let tokenResponse;
  try {
    tokenResponse = await msal.acquireTokenSilent(request);
  } catch (e) {
    if (terminal) terminal.writeLine("Acquiring token interactively...", "info");
    tokenResponse = await msal.acquireTokenPopup(request);
  }
  
  const url = `https://graph.microsoft.com/v1.0${endpoint}`;
  const headers = {
    'Authorization': `Bearer ${tokenResponse.accessToken}`,
    'Content-Type': 'application/json'
  };
  
  const options = { method, headers };
  if (payload && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
    options.body = JSON.stringify(payload);
  }
  
  try {
    const res = await fetch(url, options);
    if (!res.ok) {
      const errBody = await res.json().catch(() => ({}));
      const errMessage = errBody.error?.message || `HTTP ${res.status}: ${res.statusText}`;
      if (terminal) {
        logApiTransaction(method, `https://graph.microsoft.com/v1.0${endpoint}`, scopes.join(', '), `FAILED (${res.status})`);
        updateVisualPanel('error', { endpoint, method, message: errMessage });
        updateJsonPanel(`ERROR ${method} /v1.0${endpoint}`, { error: errMessage, status: res.status });
      }
      throw new Error(errMessage);
    }
    
    if (terminal) {
      logApiTransaction(method, `https://graph.microsoft.com/v1.0${endpoint}`, scopes.join(', '), `SUCCESS (${res.status} OK)`);
    }
    
    if (res.status === 204) return null; // No Content
    return await res.json();
  } catch (err) {
    if (terminal) {
      logApiTransaction(method, `https://graph.microsoft.com/v1.0${endpoint}`, scopes.join(', '), `FAILED (${err.message})`);
      updateVisualPanel('error', { endpoint, method, message: err.message });
      updateJsonPanel(`ERROR ${method} /v1.0${endpoint}`, { error: err.message });
    }
    throw err;
  }
}

// Commands Registry
export const commands = {
  
  // HELP COMMAND
  help: (flags, params, terminal) => {
    terminal.writeLine("GraphShell CLI Console. Available commands:", "success");
    terminal.writeLine("  help                                 - Show this help summary");
    terminal.writeLine("  login [--real]                       - Perform login (mock or Entra Real ID)");
    terminal.writeLine("  logout                               - Terminate the current session");
    terminal.writeLine("  me                                   - Retrieve active profile info");
    terminal.writeLine("  users [--list] [--create] [--delete] - Manage AD Directory Users");
    terminal.writeLine("      Flags: --list                     - List all users");
    terminal.writeLine("             --create                   - Add user (requires: -n \"Name\" -j \"Job\")");
    terminal.writeLine("             --delete                   - Remove user (requires: -id \"user-id\")");
    terminal.writeLine("  groups                               - List AD Directory Groups");
    terminal.writeLine("  mail [--list] [--send]               - View and send Outlook Emails");
    terminal.writeLine("      Flags: --list                     - Read newest emails");
    terminal.writeLine("             --send                     - Send email (requires: --to \"email\" --subject \"sub\" --body \"txt\")");
    terminal.writeLine("  calendar [--list] [--add]            - View and add calendar meetings");
    terminal.writeLine("      Flags: --list                     - Show all events");
    terminal.writeLine("             --add                      - Add event (requires: --subject \"sub\" --start \"ISO\" --end \"ISO\")");
    terminal.writeLine("  drive [--ls] [--mkdir] [--upload]    - Explore OneDrive folders and files");
    terminal.writeLine("      Flags: --ls                       - List files in folder root");
    terminal.writeLine("             --mkdir                    - Create a new directory (requires: --name \"name\")");
    terminal.writeLine("             --upload                   - Upload a mock file (requires: --name \"name\")");
    terminal.writeLine("  request [GET/POST/DELETE] [endpoint] - Execute custom Microsoft Graph REST query");
    terminal.writeLine("  ai [natural language prompt]         - Chat with Azure AI Assistant to translate commands");
    terminal.writeLine("  settings / config                    - Open developer settings panel");
    terminal.writeLine("  clear                                - Clear terminal screen");
  },

  // LOGIN COMMAND
  login: async (flags, params, terminal) => {
    terminal.writeLine("Connecting authentication service...", "info");
    const msal = getMsalInstance();
    if (!msal) {
      terminal.writeLine("Error: MSAL client is not configured. Go to Settings panel (top right) to enter Client ID.", "error");
      return;
    }
    try {
      const loginRequest = {
        scopes: ["User.Read", "User.ReadBasic.All", "Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "Files.ReadWrite", "Organization.Read.All"]
      };
      const loginResponse = await msal.loginPopup(loginRequest);
      terminal.writeLine(`Authentication SUCCESS. Logged in as: ${loginResponse.account.username}`, "success");
      
      const name = loginResponse.account.name || loginResponse.account.username;
      const connectBtn = document.getElementById('header-connect-btn');
      if (connectBtn) {
        connectBtn.textContent = name;
        connectBtn.classList.add('connected');
      }
      
      const logoutBtn = document.getElementById('header-logout-btn');
      if (logoutBtn) logoutBtn.style.display = 'inline-flex';
      
      const tenantId = loginResponse.account.homeAccountId?.split('.')[1] || 'TENANT ACTIVE';
      document.getElementById('connection-indicator').textContent = 'CONNECTED';
      document.getElementById('connection-indicator').className = 'stat-badge connection-status online';
      
      let displayTenant = localStorage.getItem('graphshell_tenant_name') || tenantId;
      if (displayTenant === '9188040d-6c67-4c5b-b112-36a304b66dad') {
        displayTenant = 'Personal Account';
      }
      document.getElementById('tenant-indicator').textContent = displayTenant;

      const terminalTitle = document.querySelector('.terminal-title');
      if (terminalTitle) {
        terminalTitle.textContent = `bash - graphshell@${loginResponse.account.username}`;
      }
      const promptSymbol = document.querySelector('.prompt-symbol');
      if (promptSymbol) {
        const simpleUser = loginResponse.account.username.split('@')[0];
        promptSymbol.textContent = `graphshell@${simpleUser}$`;
      }
      
      // Auto pull profile
      commands.me({}, [], terminal);
    } catch (e) {
      terminal.writeLine(`Authentication FAILED: ${e.message}`, "error");
    }
  },

  // LOGOUT COMMAND
  logout: async (flags, params, terminal) => {
    terminal.writeLine("Signing out...", "info");
    const msal = getMsalInstance();
    if (msal) {
      await msal.logoutPopup().catch(() => {});
    }
    
    terminal.writeLine("Session ended. All auth tokens cleared.", "success");
    localStorage.removeItem('graphshell_tenant_name');

    const connectBtn = document.getElementById('header-connect-btn');
    if (connectBtn) {
      connectBtn.textContent = 'Connect Account';
      connectBtn.classList.remove('connected');
    }

    const logoutBtn = document.getElementById('header-logout-btn');
    if (logoutBtn) logoutBtn.style.display = 'none';

    document.getElementById('connection-indicator').textContent = 'DISCONNECTED';
    document.getElementById('connection-indicator').className = 'stat-badge connection-status offline';
    document.getElementById('tenant-indicator').textContent = 'NO SESSION';

    const terminalTitle = document.querySelector('.terminal-title');
    if (terminalTitle) {
      terminalTitle.textContent = 'bash - graphshell@contoso.onmicrosoft.com';
    }
    const promptSymbol = document.querySelector('.prompt-symbol');
    if (promptSymbol) {
      promptSymbol.textContent = 'graphshell$';
    }
    
    updateVisualPanel('empty', null);
  },

  // PROFILE ME COMMAND
  me: async (flags, params, terminal) => {
    terminal.writeLine("Fetching current user profile from /v1.0/me...", "info");
    try {
      const user = await executeGraphRequest('/me', 'GET', null, terminal);
      terminal.writeLine(`Display Name:  ${user.displayName}`, "default");
      terminal.writeLine(`User Principal: ${user.userPrincipalName || user.mail}`, "default");
      terminal.writeLine(`Job Title:      ${user.jobTitle || 'N/A'}`, "default");
      terminal.writeLine(`Location:       ${user.officeLocation || 'N/A'}`, "default");
      
      updateJsonPanel('GET /v1.0/me', user);
      updateVisualPanel('me', user);

      // Update header button to show Entra ID Display Name
      const connectBtn = document.getElementById('header-connect-btn');
      if (connectBtn) {
        connectBtn.textContent = user.displayName;
        connectBtn.classList.add('connected');
      }
      
      // Update logout button
      const logoutBtn = document.getElementById('header-logout-btn');
      if (logoutBtn) {
        logoutBtn.style.display = 'inline-flex';
      }

      // Update terminal title with Entra ID user principal name
      const terminalTitle = document.querySelector('.terminal-title');
      if (terminalTitle) {
        terminalTitle.textContent = `bash - graphshell@${user.userPrincipalName || user.mail}`;
      }
      const promptSymbol = document.querySelector('.prompt-symbol');
      if (promptSymbol) {
        const simpleUser = (user.userPrincipalName || user.mail).split('@')[0];
        promptSymbol.textContent = `graphshell@${simpleUser}$`;
      }

      // Query tenant name from organization details
      try {
        const orgRes = await executeGraphRequest('/organization', 'GET', null, null);
        if (orgRes && orgRes.value && orgRes.value.length > 0) {
          const org = orgRes.value[0];
          const tenantName = org.displayName || org.id;
          document.getElementById('tenant-indicator').textContent = tenantName;
          localStorage.setItem('graphshell_tenant_name', tenantName);
        }
      } catch (orgErr) {
        console.warn("Could not fetch organization name:", orgErr);
        // Fallback check if it's a personal account GUID
        const accounts = getMsalInstance()?.getAllAccounts() || [];
        const tenantId = accounts[0]?.homeAccountId?.split('.')[1];
        if (tenantId === '9188040d-6c67-4c5b-b112-36a304b66dad') {
          document.getElementById('tenant-indicator').textContent = 'Personal Account';
          localStorage.setItem('graphshell_tenant_name', 'Personal Account');
        }
      }
    } catch (e) {
      terminal.writeLine(`Graph Query Error: ${e.message}`, "error");
    }
  },

  // DIRECTORY USERS COMMAND
  users: async (flags, params, terminal) => {
    // CREATE USER
    if (flags.create) {
      const name = flags.n || flags.name;
      const job = flags.j || flags.jobTitle;
      const dept = flags.d || flags.department;
      
      if (!name) {
        terminal.writeLine("Error: Missing parameters. Usage: users --create --name \"John\" --jobTitle \"Engineer\" --department \"Tech\"", "error");
        return;
      }
      
      terminal.writeLine(`Posting new user registration payload to /v1.0/users...`, "info");
      
      const payload = {
        accountEnabled: true,
        displayName: name,
        mailNickname: name.toLowerCase().replace(/\s+/g, ''),
        userPrincipalName: `${name.toLowerCase().replace(/\s+/g, '')}@contoso.onmicrosoft.com`,
        jobTitle: job || "Contributor",
        department: dept || "General",
        officeLocation: "Seattle HQ"
      };

      try {
        const created = await executeGraphRequest('/users', 'POST', payload, terminal);
        terminal.writeLine(`User created successfully:`, "success");
        terminal.writeLine(`  ID:           ${created.id}`);
        terminal.writeLine(`  UPN:          ${created.userPrincipalName}`);
        terminal.writeLine(`  Department:   ${created.department}`);
        
        updateJsonPanel('POST /v1.0/users', created);
        // Refresh grid
        commands.users({ list: true }, [], terminal);
      } catch (e) {
        terminal.writeLine(`Failed to create user: ${e.message}`, "error");
      }
      return;
    }

    // DELETE USER
    if (flags.delete) {
      const id = flags.id;
      if (!id) {
        terminal.writeLine("Error: Missing --id argument. Usage: users --delete --id \"user-id-or-name\"", "error");
        return;
      }
      
      terminal.writeLine(`Sending DELETE request to /v1.0/users/${id}...`, "info");
      try {
        const res = await executeGraphRequest(`/users/${id}`, 'DELETE', null, terminal);
        terminal.writeLine(`User resource deleted successfully.`, "success");
        
        updateJsonPanel(`DELETE /v1.0/users/${id}`, res);
        // Refresh grid
        commands.users({ list: true }, [], terminal);
      } catch (e) {
        terminal.writeLine(`Failed to delete user: ${e.message}`, "error");
      }
      return;
    }

    // LIST USERS (default or explicit)
    terminal.writeLine("Querying Azure Active Directory users from /v1.0/users...", "info");
    try {
      // Build filters if present
      let endpoint = '/users';
      if (flags.department) {
        endpoint += `?$filter=department eq '${flags.department}'`;
      }
      
      const res = await executeGraphRequest(endpoint, 'GET', null, terminal);
      const userList = res.value || [];
      
      terminal.writeLine(`Query returned ${userList.length} users:`, "success");
      userList.forEach(u => {
        terminal.writeLine(`  [${u.id.substring(0, 8)}...] ${u.displayName} - ${u.jobTitle || 'Member'} (${u.department || 'General'})`);
      });
      
      updateJsonPanel(`GET /v1.0${endpoint}`, res);
      updateVisualPanel('users', userList);
    } catch (e) {
      terminal.writeLine(`Graph Query Error: ${e.message}`, "error");
    }
  },

  // GROUPS COMMAND
  groups: async (flags, params, terminal) => {
    terminal.writeLine("Querying Active Directory groups from /v1.0/groups...", "info");
    try {
      const res = await executeGraphRequest('/groups', 'GET', null, terminal);
      const groupList = res.value || [];
      
      terminal.writeLine(`Query returned ${groupList.length} groups:`, "success");
      groupList.forEach(g => {
        terminal.writeLine(`  [${g.id.substring(0, 8)}...] ${g.displayName} - ${g.description || 'No Description'}`);
      });
      
      updateJsonPanel('GET /v1.0/groups', res);
      updateVisualPanel('groups', groupList);
    } catch (e) {
      terminal.writeLine(`Graph Query Error: ${e.message}`, "error");
    }
  },

  // MAIL COMMAND
  mail: async (flags, params, terminal) => {
    // SEND MAIL
    if (flags.send) {
      const to = flags.to;
      const sub = flags.subject || flags.sub;
      const body = flags.body;
      
      if (!to || !body) {
        terminal.writeLine("Error: Missing parameters. Usage: mail --send --to \"alex@contoso.com\" --subject \"Update\" --body \"Sync at 2\"", "error");
        return;
      }
      
      terminal.writeLine(`Posting email payload to /v1.0/me/sendMail...`, "info");
      const payload = {
        message: {
          subject: sub || "(No Subject)",
          body: {
            contentType: "Text",
            content: body
          },
          toRecipients: [
            {
              emailAddress: {
                address: to
              }
            }
          ]
        }
      };

      try {
        const sent = await executeGraphRequest('/me/sendMail', 'POST', payload, terminal);
        terminal.writeLine(`Email dispatched successfully to ${to}.`, "success");
        updateJsonPanel('POST /v1.0/me/sendMail', sent);
        // Refresh inbox
        commands.mail({ list: true }, [], terminal);
      } catch (e) {
        terminal.writeLine(`Failed to send email: ${e.message}`, "error");
      }
      return;
    }

    // LIST MAIL (default or explicit)
    terminal.writeLine("Querying Outlook inbox from /v1.0/me/messages...", "info");
    try {
      const res = await executeGraphRequest('/me/messages', 'GET', null, terminal);
      const messages = res.value || [];
      
      terminal.writeLine(`Inbox contains ${messages.length} messages:`, "success");
      messages.forEach(m => {
        const sender = m.from?.emailAddress?.address || m.from || "System";
        terminal.writeLine(`  [${sender}] ${m.subject || '(No Subject)'}`);
      });
      
      updateJsonPanel('GET /v1.0/me/messages', res);
      updateVisualPanel('mail', messages);
    } catch (e) {
      terminal.writeLine(`Graph Query Error: ${e.message}`, "error");
    }
  },

  // CALENDAR COMMAND
  calendar: async (flags, params, terminal) => {
    // ADD EVENT
    if (flags.add) {
      const sub = flags.subject || flags.sub;
      const start = flags.start;
      const end = flags.end;
      const loc = flags.location || flags.loc || "Teams Meeting";
      
      if (!sub || !start || !end) {
        terminal.writeLine("Error: Missing parameters. Usage: calendar --add --subject \"Sync\" --start \"2026-05-21T14:00\" --end \"2026-05-21T14:30\"", "error");
        return;
      }
      
      terminal.writeLine(`Scheduling calendar event in /v1.0/me/events...`, "info");
      const payload = {
        subject: sub,
        start: {
          dateTime: start,
          timeZone: "Pacific Standard Time"
        },
        end: {
          dateTime: end,
          timeZone: "Pacific Standard Time"
        },
        location: {
          displayName: loc
        }
      };

      try {
        const added = await executeGraphRequest('/me/events', 'POST', payload, terminal);
        terminal.writeLine(`Meeting added to calendar successfully.`, "success");
        updateJsonPanel('POST /v1.0/me/events', added);
        // Refresh calendar view
        commands.calendar({ list: true }, [], terminal);
      } catch (e) {
        terminal.writeLine(`Failed to schedule event: ${e.message}`, "error");
      }
      return;
    }

    // LIST CALENDAR EVENTS (default or explicit)
    terminal.writeLine("Fetching calendar timeline from /v1.0/me/events...", "info");
    try {
      const res = await executeGraphRequest('/me/events', 'GET', null, terminal);
      const events = res.value || [];
      
      terminal.writeLine(`Found ${events.length} upcoming appointments:`, "success");
      events.forEach(ev => {
        const time = new Date(ev.start?.dateTime || ev.start).toLocaleString();
        terminal.writeLine(`  [${time}] ${ev.subject} @ ${ev.location?.displayName || 'Online'}`);
      });
      
      updateJsonPanel('GET /v1.0/me/events', res);
      updateVisualPanel('calendar', events);
    } catch (e) {
      terminal.writeLine(`Graph Query Error: ${e.message}`, "error");
    }
  },

  // DRIVE COMMAND
  drive: async (flags, params, terminal) => {
    const parentFolder = '/';
    
    // MAKE DIRECTORY (FOLDER)
    if (flags.mkdir) {
      const name = flags.name || flags.n;
      if (!name) {
        terminal.writeLine("Error: Missing --name argument. Usage: drive --mkdir --name \"NewFolder\"", "error");
        return;
      }
      terminal.writeLine(`Creating new directory '${name}' in /v1.0/me/drive/root/children...`, "info");
      const payload = {
        name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename"
      };
      try {
        const folder = await executeGraphRequest('/me/drive/root/children', 'POST', payload, terminal);
        terminal.writeLine(`Directory created: ${folder.name}`, "success");
        updateJsonPanel('POST /v1.0/me/drive/root/children', folder);
        // Refresh drive
        commands.drive({ ls: true }, [], terminal);
      } catch (e) {
        terminal.writeLine(`Failed to create directory: ${e.message}`, "error");
      }
      return;
    }

    // UPLOAD MOCK FILE
    if (flags.upload) {
      const name = flags.name || flags.n;
      if (!name) {
        terminal.writeLine("Error: Missing --name argument. Usage: drive --upload --name \"report.txt\"", "error");
        return;
      }
      terminal.writeLine(`Uploading file content stream to /v1.0/me/drive/root/children/${name}/content...`, "info");
      const payload = {
        name,
        size: '22 KB'
      };
      try {
        const file = await executeGraphRequest('/me/drive/root/children', 'POST', payload, terminal);
        terminal.writeLine(`File uploaded successfully: ${file.name} (${file.size})`, "success");
        updateJsonPanel('POST /v1.0/me/drive/root/children', file);
        // Refresh drive
        commands.drive({ ls: true }, [], terminal);
      } catch (e) {
        terminal.writeLine(`Failed to upload file: ${e.message}`, "error");
      }
      return;
    }

    // LS (LIST FILES)
    terminal.writeLine(`Listing folder resources in OneDrive root: /me/drive/root/children...`, "info");
    try {
      const res = await executeGraphRequest('/me/drive/root/children', 'GET', null, terminal);
      const items = res.value || [];
      
      terminal.writeLine(`Directory root contains ${items.length} nodes:`, "success");
      items.forEach(it => {
        const type = it.folder ? '<DIR>' : '<FILE>';
        terminal.writeLine(`  ${type.padEnd(8)} ${it.name.padEnd(25)} ${it.size || '--'}`);
      });
      
      updateJsonPanel('GET /v1.0/me/drive/root/children', res);
      updateVisualPanel('drive', items);
    } catch (e) {
      terminal.writeLine(`Graph Query Error: ${e.message}`, "error");
    }
  },

  // RAW REQUEST (GRAPH EXPLORER) COMMAND
  request: async (flags, params, terminal) => {
    const method = params[0]?.toUpperCase() || 'GET';
    const endpoint = params[1];
    
    if (!endpoint) {
      terminal.writeLine("Error: Missing parameters. Usage: request [GET/POST/DELETE] [endpoint_path]", "error");
      terminal.writeLine("Example: request GET /me/messages", "info");
      return;
    }
    
    terminal.writeLine(`Sending raw HTTP ${method} to https://graph.microsoft.com/v1.0${endpoint}...`, "info");
    try {
      // For POST/PUT requests, we can look for a --body flag or prompt
      let payload = null;
      if (flags.body) {
        try {
          payload = JSON.parse(flags.body);
        } catch (err) {
          throw new Error("Invalid JSON in --body flag.");
        }
      }
      
      const res = await executeGraphRequest(endpoint, method, payload, terminal);
      terminal.writeLine("Request finished with success. JSON payload logged.", "success");
      
      updateJsonPanel(`${method} /v1.0${endpoint}`, res);
      
      // Determine a visual tab fit
      if (endpoint.includes('/messages')) {
        updateVisualPanel('mail', res.value || [res]);
      } else if (endpoint.includes('/users')) {
        updateVisualPanel('users', res.value || [res]);
      } else if (endpoint.includes('/events')) {
        updateVisualPanel('calendar', res.value || [res]);
      } else {
        updateVisualPanel('raw', res);
      }
    } catch (e) {
      terminal.writeLine(`HTTP Request Failed: ${e.message}`, "error");
    }
  },

  // AI COPILOT COMMAND
  ai: async (flags, params, terminal) => {
    const prompt = params.join(' ');
    if (!prompt) {
      terminal.writeLine("Error: Please provide a prompt for the AI. Usage: ai [your request in plain English]", "error");
      return;
    }
    
    // Switch to AI tab visually
    document.getElementById('tab-ai').click();
    
    terminal.writeLine(`Sending request to Azure AI Service: "${prompt}"...`, "info");
    
    const response = await runAiPrompt(prompt, terminal);
    
    terminal.writeLine(`Azure AI Assistant Response:`, "success");
    terminal.writeLine(response.explanation, "default");
    if (response.command) {
      terminal.writeLine(`Mapped Command: ${response.command}`, "info");
      // Populate command to input field for easy execution
      document.getElementById('terminal-input').value = response.command;
      terminal.writeLine(`(The command is pre-filled in your prompt. Press [Enter] to run it)`, "info");
    }
  },

  // CONFIG COMMAND
  config: (flags, params, terminal) => {
    terminal.writeLine("Opening system configuration panel...", "success");
    const openSettingsBtn = document.getElementById('open-settings-btn');
    if (openSettingsBtn) openSettingsBtn.click();
  },

  // SETTINGS COMMAND
  settings: (flags, params, terminal) => {
    terminal.writeLine("Opening system configuration panel...", "success");
    const openSettingsBtn = document.getElementById('open-settings-btn');
    if (openSettingsBtn) openSettingsBtn.click();
  }

};
