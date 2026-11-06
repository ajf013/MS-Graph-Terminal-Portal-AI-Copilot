import { TerminalConsole } from './terminal.js';
import { commands, registerCallbacks } from './commands.js';
import { runAiPrompt } from './aiService.js';

const DEFAULT_CLIENT_ID = '8814eadb-592a-4da9-be40-94a8fb602420';
const DEFAULT_TENANT_ID = 'common';
const DEFAULT_AI_ENDPOINT = 'https://cog-graphshell-ai-4a27a.openai.azure.com/';
const DEFAULT_AI_DEPLOYMENT = 'gpt-5-mini';
const DEFAULT_AI_KEY = '';

// MSAL Instance reference
let msalInstance = null;

// DOM Elements
const connectionIndicator = document.getElementById('connection-indicator');
const tenantIndicator = document.getElementById('tenant-indicator');
const clockDisplay = document.getElementById('clock-display');
const openSettingsBtn = document.getElementById('open-settings-btn');
const closeSettingsBtn = document.getElementById('close-settings-btn');
const cancelSettingsBtn = document.getElementById('cancel-settings-btn');
const saveSettingsBtn = document.getElementById('save-settings-btn');
const resetDbBtn = document.getElementById('reset-db-btn');
const settingsModal = document.getElementById('settings-modal');
const msalFieldsContainer = document.getElementById('msal-fields-container');
const headerConnectBtn = document.getElementById('header-connect-btn');

// Settings Inputs
const msalClientIdInput = document.getElementById('msal-client-id');
const msalTenantIdInput = document.getElementById('msal-tenant-id');
const aiEndpointInput = document.getElementById('ai-endpoint');
const aiDeploymentInput = document.getElementById('ai-deployment');
const aiApiKeyInput = document.getElementById('ai-api-key');

// Tabs & Content
const tabBtns = document.querySelectorAll('.tab-btn');
const tabPanels = document.querySelectorAll('.tab-panel');
const visualContentArea = document.getElementById('visual-content-area');
const jsonOutputArea = document.getElementById('json-output-area');
const jsonEndpointDisplay = document.getElementById('json-endpoint-display');
const copyJsonBtn = document.getElementById('copy-json-btn');

// AI Chat Elements
const aiChatLog = document.getElementById('ai-chat-log');
const aiChatInput = document.getElementById('ai-chat-input');
const aiChatSendBtn = document.getElementById('ai-chat-send-btn');

// Terminal Reference
let terminalConsole = null;

// Current state variables
let currentPath = '/';

/* ==========================================================================
   INITIALIZATION
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {
  // PWA Service Worker Registration
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js')
      .then(() => console.log('Service Worker Registered.'))
      .catch(err => console.warn('Service Worker Registration Failed:', err));
  }

  // Start Clock ticking
  updateClock();
  setInterval(updateClock, 1000);

  // Load Saved Configurations
  loadSettings();

  // Initialize MSAL
  initMsal();

  // Create Terminal instance
  const terminalInput = document.getElementById('terminal-input');
  const terminalHistoryLog = document.getElementById('terminal-history-log');
  terminalConsole = new TerminalConsole(terminalInput, terminalHistoryLog, commands);

  // Register command callbacks to update dashboard components
  registerCallbacks(
    renderVisualView,
    renderJsonView,
    () => msalInstance,
    updateApiLoggerWidget
  );

  // Setup Event Listeners
  initUIEvents();

  // If already logged in, automatically run profile fetch to update UI elements
  if (msalInstance && msalInstance.getAllAccounts().length > 0) {
    setTimeout(() => {
      terminalConsole.writeLine("Restoring active session from cache...", "info");
      terminalConsole.handleCommand('me');
    }, 300);
  }
});

/* ==========================================================================
   SETTINGS & AUTHENTICATION
   ========================================================================== */

function loadSettings() {
  msalClientIdInput.value = localStorage.getItem('graphshell_msal_client_id') || DEFAULT_CLIENT_ID;
  msalTenantIdInput.value = localStorage.getItem('graphshell_msal_tenant_id') || DEFAULT_TENANT_ID;
  aiEndpointInput.value = localStorage.getItem('graphshell_ai_endpoint') || DEFAULT_AI_ENDPOINT;
  aiDeploymentInput.value = localStorage.getItem('graphshell_ai_deployment') || DEFAULT_AI_DEPLOYMENT;
  aiApiKeyInput.value = localStorage.getItem('graphshell_ai_key') || DEFAULT_AI_KEY;
}

function saveSettings() {
  localStorage.setItem('graphshell_msal_client_id', msalClientIdInput.value.trim() || DEFAULT_CLIENT_ID);
  localStorage.setItem('graphshell_msal_tenant_id', msalTenantIdInput.value.trim() || DEFAULT_TENANT_ID);
  localStorage.setItem('graphshell_ai_endpoint', aiEndpointInput.value.trim() || DEFAULT_AI_ENDPOINT);
  localStorage.setItem('graphshell_ai_deployment', aiDeploymentInput.value.trim() || DEFAULT_AI_DEPLOYMENT);
  localStorage.setItem('graphshell_ai_key', aiApiKeyInput.value.trim() || DEFAULT_AI_KEY);

  settingsModal.classList.remove('active');
  
  // Re-load app settings state
  loadSettings();
  initMsal();

  if (terminalConsole) {
    terminalConsole.writeLine("System configurations updated. Resetting shell session.", "warning");
    terminalConsole.writeLine("Please use 'login' to sign in to your real Microsoft account.", "info");
  }
}

function initMsal() {
  const clientId = localStorage.getItem('graphshell_msal_client_id') || DEFAULT_CLIENT_ID;
  const tenantId = localStorage.getItem('graphshell_msal_tenant_id') || DEFAULT_TENANT_ID;

  if (clientId) {
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri: window.location.href.split('?')[0]
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
      }
    };
    try {
      msalInstance = new msal.PublicClientApplication(msalConfig);
      
      const accounts = msalInstance.getAllAccounts();
      const connectBtn = document.getElementById('header-connect-btn');
      const logoutBtn = document.getElementById('header-logout-btn');
      if (accounts.length > 0) {
        const name = accounts[0].name || accounts[0].username;
        connectBtn.textContent = name;
        connectBtn.classList.add('connected');
        if (logoutBtn) logoutBtn.style.display = 'inline-flex';
        
        const homeTenant = accounts[0].homeAccountId?.split('.')[1] || 'TENANT ACTIVE';
        connectionIndicator.textContent = 'CONNECTED';
        connectionIndicator.className = 'stat-badge connection-status online';
        
        let displayTenant = localStorage.getItem('graphshell_tenant_name') || homeTenant;
        if (displayTenant === '9188040d-6c67-4c5b-b112-36a304b66dad') {
          displayTenant = 'Personal Account';
        }
        tenantIndicator.textContent = displayTenant;

        const terminalTitle = document.querySelector('.terminal-title');
        if (terminalTitle) {
          terminalTitle.textContent = `bash - graphshell@${accounts[0].username}`;
        }
        const promptSymbol = document.querySelector('.prompt-symbol');
        if (promptSymbol) {
          const simpleUser = accounts[0].username.split('@')[0];
          promptSymbol.textContent = `graphshell@${simpleUser}$`;
        }
      } else {
        connectBtn.textContent = 'Connect Account';
        connectBtn.classList.remove('connected');
        if (logoutBtn) logoutBtn.style.display = 'none';
        connectionIndicator.textContent = 'DISCONNECTED';
        connectionIndicator.className = 'stat-badge connection-status offline';
        tenantIndicator.textContent = 'NO SESSION';

        const terminalTitle = document.querySelector('.terminal-title');
        if (terminalTitle) {
          terminalTitle.textContent = 'bash - graphshell@contoso.onmicrosoft.com';
        }
        const promptSymbol = document.querySelector('.prompt-symbol');
        if (promptSymbol) {
          promptSymbol.textContent = 'graphshell$';
        }
      }
    } catch (err) {
      console.error("MSAL init failure:", err);
    }
  } else {
    msalInstance = null;
  }
}

/* ==========================================================================
   UI STATE & TABS
   ========================================================================== */

function initUIEvents() {
  // Settings Dialog handlers
  openSettingsBtn.addEventListener('click', () => settingsModal.classList.add('active'));
  closeSettingsBtn.addEventListener('click', () => settingsModal.classList.remove('active'));
  cancelSettingsBtn.addEventListener('click', () => settingsModal.classList.remove('active'));
  saveSettingsBtn.addEventListener('click', saveSettings);
  
  // Connect Account button handler
  headerConnectBtn.addEventListener('click', () => {
    const accounts = msalInstance ? msalInstance.getAllAccounts() : [];
    if (accounts.length > 0) {
      terminalConsole.handleCommand('me');
    } else {
      terminalConsole.handleCommand('login');
    }
  });

  // Logout button handler
  const headerLogoutBtn = document.getElementById('header-logout-btn');
  if (headerLogoutBtn) {
    headerLogoutBtn.addEventListener('click', () => {
      terminalConsole.handleCommand('logout');
    });
  }

  // Learn tab interactive buttons
  document.querySelectorAll('.try-cmd-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const cmd = btn.getAttribute('data-cmd');
      if (terminalConsole) {
        const input = document.getElementById('terminal-input');
        input.value = cmd;
        input.focus();
        terminalConsole.handleCommand(cmd);
      }
    });
  });

  resetDbBtn.addEventListener('click', () => {
    if (confirm("This will clear your credentials and saved command history. Continue?")) {
      localStorage.clear();
      if (terminalConsole) {
        terminalConsole.history = [];
        terminalConsole.historyIndex = 0;
        terminalConsole.clear();
        terminalConsole.writeLine("Configurations and command history flushed.", "success");
      }
      settingsModal.classList.remove('active');
      loadSettings();
      initMsal();
      renderVisualView('empty', null);
    }
  });

  // Dynamic panel tabs switching
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      tabPanels.forEach(p => p.classList.remove('active'));
      
      btn.classList.add('active');
      const targetPanel = document.getElementById(btn.getAttribute('data-target'));
      targetPanel.classList.add('active');
    });
  });

  // Copy JSON code payload
  copyJsonBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(jsonOutputArea.textContent)
      .then(() => {
        copyJsonBtn.textContent = 'Copied!';
        setTimeout(() => copyJsonBtn.textContent = 'Copy Payload', 1500);
      });
  });

  // AI Chat inputs and click links
  aiChatSendBtn.addEventListener('click', handleAiMessage);
  aiChatInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') handleAiMessage();
  });

  document.querySelectorAll('.ai-suggestion').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      aiChatInput.value = e.target.textContent.replace(/"/g, '');
      handleAiMessage();
    });
  });
}

function updateClock() {
  const now = new Date();
  clockDisplay.textContent = now.toTimeString().split(' ')[0];
}

/* ==========================================================================
   AI COPILOT CONTROLLER
   ========================================================================== */

async function handleAiMessage() {
  const text = aiChatInput.value.trim();
  if (!text) return;
  aiChatInput.value = '';

  // Append user message bubble
  appendChatBubble('user', text);

  // Append loading animation bubble
  const loaderId = appendChatLoadingBubble();

  // Send request to AI Service
  try {
    const res = await runAiPrompt(text, terminalConsole);
    removeChatLoadingBubble(loaderId);
    
    // Create assistant response content
    const wrapper = document.createElement('div');
    const p1 = document.createElement('p');
    p1.textContent = res.explanation;
    wrapper.appendChild(p1);

    if (res.command) {
      const p2 = document.createElement('p');
      p2.innerHTML = `<strong>Suggested CLI action:</strong>`;
      wrapper.appendChild(p2);

      const codeLink = document.createElement('a');
      codeLink.className = 'cmd-shortcut';
      codeLink.href = '#';
      codeLink.style.display = 'block';
      codeLink.style.marginTop = '4px';
      codeLink.textContent = res.command;
      codeLink.addEventListener('click', (e) => {
        e.preventDefault();
        const input = document.getElementById('terminal-input');
        input.value = res.command;
        input.focus();
      });
      wrapper.appendChild(codeLink);
    }
    
    appendChatBubble('assistant', wrapper);
  } catch (err) {
    removeChatLoadingBubble(loaderId);
    appendChatBubble('assistant', `Failed to retrieve response: ${err.message}`);
  }
}

function appendChatBubble(sender, content) {
  const msg = document.createElement('div');
  msg.className = `ai-msg ${sender}`;
  
  const avatar = document.createElement('div');
  avatar.className = 'msg-avatar';
  avatar.textContent = sender === 'user' ? 'U' : 'AI';
  
  const bubble = document.createElement('div');
  bubble.className = 'msg-content';
  
  if (typeof content === 'string') {
    bubble.textContent = content;
  } else {
    bubble.appendChild(content);
  }
  
  msg.appendChild(avatar);
  msg.appendChild(bubble);
  aiChatLog.appendChild(msg);
  aiChatLog.scrollTop = aiChatLog.scrollHeight;
}

function appendChatLoadingBubble() {
  const id = `loading-${Date.now()}`;
  const msg = document.createElement('div');
  msg.className = 'ai-msg assistant';
  msg.id = id;
  
  const avatar = document.createElement('div');
  avatar.className = 'msg-avatar';
  avatar.textContent = 'AI';
  
  const bubble = document.createElement('div');
  bubble.className = 'msg-content';
  bubble.innerHTML = '<span class="pulse-icon">Analyzing schema parameters...</span>';
  
  msg.appendChild(avatar);
  msg.appendChild(bubble);
  aiChatLog.appendChild(msg);
  aiChatLog.scrollTop = aiChatLog.scrollHeight;
  return id;
}

function removeChatLoadingBubble(id) {
  const loader = document.getElementById(id);
  if (loader) loader.remove();
}

function updateApiLoggerWidget(method, url, scopes, status) {
  const statusEl = document.getElementById('api-log-status');
  const methodEl = document.getElementById('api-log-method');
  const urlEl = document.getElementById('api-log-url');
  const scopesEl = document.getElementById('api-log-scopes');
  const docEl = document.getElementById('api-log-doc');
  
  if (statusEl) {
    statusEl.textContent = status;
    statusEl.className = 'api-logger-status';
    if (status.startsWith('SUCCESS')) {
      statusEl.classList.add('success');
    } else if (status.startsWith('FAILED')) {
      statusEl.classList.add('failed');
    }
  }
  
  if (methodEl) methodEl.textContent = method;
  if (urlEl) urlEl.textContent = url;
  if (scopesEl) scopesEl.textContent = scopes;
  
  // Set documentation link based on URL
  if (docEl) {
    let docUrl = "https://learn.microsoft.com/graph/api/overview";
    if (url.includes('/me/messages')) {
      docUrl = "https://learn.microsoft.com/graph/api/user-list-messages";
    } else if (url.includes('/me/events')) {
      docUrl = "https://learn.microsoft.com/graph/api/user-list-events";
    } else if (url.includes('/me/drive')) {
      docUrl = "https://learn.microsoft.com/graph/api/drive-get";
    } else if (url.includes('/users')) {
      docUrl = "https://learn.microsoft.com/graph/api/user-list";
    } else if (url.includes('/groups')) {
      docUrl = "https://learn.microsoft.com/graph/api/group-list";
    } else if (url.endsWith('/me')) {
      docUrl = "https://learn.microsoft.com/graph/api/user-get";
    }
    docEl.href = docUrl;
  }
}

/* ==========================================================================
   RENDER EXPLORER PANELS
   ========================================================================== */

function renderJsonView(endpoint, data) {
  jsonEndpointDisplay.textContent = endpoint;
  jsonOutputArea.textContent = JSON.stringify(data, null, 2);
}

// Route rendering requests to dynamic widgets
function renderVisualView(viewName, data) {
  visualContentArea.innerHTML = '';
  
  if (viewName === 'empty') {
    visualContentArea.innerHTML = `
      <div class="empty-state">
        <img src="logo.png" alt="GraphShell Logo" class="pulse-icon" style="width: 64px; height: 64px; border-radius: 12px; box-shadow: 0 0 20px rgba(0, 240, 255, 0.2); margin-bottom: 16px;">
        <h3>Waiting for Command</h3>
        <p>Run a directory command in the shell to visualize results here. Try <code class="cmd-shortcut">me</code>, <code class="cmd-shortcut">users --list</code>, or <code class="cmd-shortcut">mail --list</code>.</p>
      </div>
    `;
    return;
  }

  if (viewName === 'me') {
    renderProfileCard(data);
    return;
  }

  if (viewName === 'error') {
    renderErrorCard(data);
    return;
  }

  if (viewName === 'users') {
    renderUsersGrid(data);
    return;
  }

  if (viewName === 'groups') {
    renderGroupsList(data);
    return;
  }

  if (viewName === 'mail') {
    renderMailApp(data);
    return;
  }

  if (viewName === 'calendar') {
    renderCalendarApp(data);
    return;
  }

  if (viewName === 'drive') {
    renderDriveApp(data);
    return;
  }

  // Fallback raw object visualizer
  const container = document.createElement('div');
  container.style.width = '100%';
  container.innerHTML = `<h3>Result Payload</h3><p>Visual mapping not registered for this endpoint. Inspect the JSON tab for details.</p>`;
  visualContentArea.appendChild(container);
}

/* ==========================================================================
   DASHBOARD COMPONENT RENDERERS
   ========================================================================== */

// Helper to acquire a silent Graph token
async function getGraphToken(scopes = ["User.Read"]) {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return null;
  try {
    const res = await msalInstance.acquireTokenSilent({
      scopes,
      account: accounts[0]
    });
    return res.accessToken;
  } catch (err) {
    console.warn("Silent token acquisition failed:", err);
    return null;
  }
}

// Helper to fetch user profile photo blob URL
async function fetchUserProfilePhoto(token) {
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });
    if (res.ok) {
      const blob = await res.blob();
      return URL.createObjectURL(blob);
    }
  } catch (err) {
    console.warn("Could not fetch user photo:", err);
  }
  return null;
}

function renderProfileCard(user) {
  const initials = user.displayName ? user.displayName.split(' ').map(n => n[0]).join('') : 'U';
  const card = document.createElement('div');
  card.className = 'user-view';
  card.style.display = 'flex';
  card.style.flexDirection = 'column';
  card.style.alignItems = 'center';
  card.style.gap = '14px';
  card.style.width = '100%';
  card.style.padding = '20px';

  card.innerHTML = `
    <div class="user-avatar" id="user-profile-photo-container" style="width: 80px; height: 80px; font-size: 2rem; border-radius: 50%; display: flex; align-items: center; justify-content: center; background-size: cover; background-position: center; box-shadow: var(--shadow-neon);">${initials}</div>
    <h2 style="font-size: 1.4rem;">${user.displayName}</h2>
    <div style="font-size: 0.85rem; color: var(--accent-cyan); font-family: var(--font-mono);">${user.jobTitle || 'User'}</div>
    
    <div style="width: 100%; border-top: var(--border-dim); margin-top: 10px; padding-top: 16px; display: flex; flex-direction: column; gap: 8px;">
      <div><span style="color: var(--text-dim)">UPN/Email:</span> <span style="font-family: var(--font-mono)">${user.userPrincipalName || user.mail}</span></div>
      <div><span style="color: var(--text-dim)">Office:</span> <span>${user.officeLocation || 'Remote'}</span></div>
      <div><span style="color: var(--text-dim)">ID:</span> <span style="font-family: var(--font-mono); font-size: 0.8rem; color: var(--accent-amber);">${user.id}</span></div>
    </div>
  `;
  visualContentArea.appendChild(card);

  // Fetch photo and apply if found
  getGraphToken(["User.Read"]).then(token => {
    if (token) {
      fetchUserProfilePhoto(token).then(photoUrl => {
        if (photoUrl) {
          const imgContainer = document.getElementById('user-profile-photo-container');
          if (imgContainer) {
            imgContainer.style.backgroundImage = `url(${photoUrl})`;
            imgContainer.textContent = ''; // Clear initials
          }
        }
      });
    }
  });
}

function renderUsersGrid(users) {
  const container = document.createElement('div');
  container.style.width = '100%';
  container.innerHTML = `
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
      <h3 style="font-size: 1.1rem;">Directory Active Users</h3>
      <span class="stat-badge" style="font-family: var(--font-mono);">${users.length} Users</span>
    </div>
  `;

  const grid = document.createElement('div');
  grid.className = 'users-grid';

  users.forEach(u => {
    const card = document.createElement('div');
    card.className = 'user-card';
    const initials = u.displayName.split(' ').map(n => n[0]).join('');

    card.innerHTML = `
      <div class="user-avatar">${initials}</div>
      <div class="user-name" title="${u.displayName}">${u.displayName}</div>
      <div class="user-title" title="${u.jobTitle}">${u.jobTitle}</div>
      <div class="user-badge">${u.department}</div>
      <button class="delete-user-btn" style="position: absolute; top: 6px; right: 6px; background: none; border: none; color: var(--text-dim); cursor: pointer; font-size: 0.85rem;" title="Delete User">&times;</button>
    `;

    // Click to copy delete command to CLI
    card.querySelector('.delete-user-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      if (confirm(`Do you want to write the command to delete ${u.displayName}?`)) {
        const input = document.getElementById('terminal-input');
        input.value = `users --delete -id "${u.id}"`;
        input.focus();
      }
    });

    grid.appendChild(card);
  });

  container.appendChild(grid);
  visualContentArea.appendChild(container);
}

function renderGroupsList(groups) {
  const container = document.createElement('div');
  container.style.width = '100%';
  container.innerHTML = `
    <h3 style="margin-bottom: 16px; font-size: 1.1rem;">Security & Distribution Groups</h3>
    <div style="display: flex; flex-direction: column; gap: 8px;" id="groups-list-rows"></div>
  `;

  const rows = container.querySelector('#groups-list-rows');
  groups.forEach(g => {
    const item = document.createElement('div');
    item.className = 'drive-item';
    item.style.cursor = 'default';
    item.innerHTML = `
      <div class="drive-item-meta">
        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color: var(--accent-amber);"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></circle><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
        <div>
          <div class="drive-name">${g.displayName}</div>
          <div style="font-size: 0.75rem; color: var(--text-dim);">${g.description}</div>
        </div>
      </div>
      <div style="font-family: var(--font-mono); font-size: 0.7rem; color: var(--text-dim);">${g.mailEnabled ? 'MAIL-ENABLED' : 'SECURITY'}</div>
    `;
    rows.appendChild(item);
  });

  visualContentArea.appendChild(container);
}

function renderMailApp(messages) {
  const mailApp = document.createElement('div');
  mailApp.className = 'mail-app';
  
  const toolbar = document.createElement('div');
  toolbar.className = 'mail-toolbar';
  toolbar.innerHTML = `
    <h3 style="font-size: 1.1rem;">Outlook Inbox</h3>
    <button class="btn-compose">
      <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 1 1 3 3L12 15l-4 1 1-4z"></path></svg>
      Compose Mail
    </button>
  `;

  const list = document.createElement('div');
  list.className = 'mail-list';

  if (messages.length === 0) {
    list.innerHTML = `<div class="empty-state"><p>Inbox is empty.</p></div>`;
  } else {
    messages.forEach(m => {
      const item = document.createElement('div');
      item.className = 'mail-item';
      
      const sender = m.from?.emailAddress?.address || m.from || "System";
      const dateStr = new Date(m.date || m.receivedDateTime).toLocaleDateString();

      item.innerHTML = `
        <div class="mail-item-header">
          <span class="mail-sender">${sender}</span>
          <span class="mail-date">${dateStr}</span>
        </div>
        <div class="mail-subject">${m.subject || '(No Subject)'}</div>
        <div class="mail-preview">${m.body?.content || m.body || ''}</div>
      `;

      item.addEventListener('click', () => renderMailItemDetail(m, messages));
      list.appendChild(item);
    });
  }

  toolbar.querySelector('.btn-compose').addEventListener('click', () => renderMailComposeForm(messages));

  mailApp.appendChild(toolbar);
  mailApp.appendChild(list);
  visualContentArea.appendChild(mailApp);
}

function renderMailItemDetail(mail, messages) {
  visualContentArea.innerHTML = '';
  const view = document.createElement('div');
  view.className = 'mail-view';

  const backBtn = document.createElement('button');
  backBtn.className = 'mail-view-back';
  backBtn.innerHTML = `&larr; Back to Inbox`;
  backBtn.addEventListener('click', () => {
    visualContentArea.innerHTML = '';
    renderMailApp(messages);
  });

  const sender = mail.from?.emailAddress?.address || mail.from || "System";
  const dateStr = new Date(mail.date || mail.receivedDateTime).toLocaleString();

  view.appendChild(backBtn);
  
  const header = document.createElement('div');
  header.innerHTML = `
    <h2 class="mail-view-title">${mail.subject || '(No Subject)'}</h2>
    <div class="mail-view-meta">
      <div><strong>From:</strong> ${sender}</div>
      <div><strong>Date:</strong> ${dateStr}</div>
    </div>
  `;
  view.appendChild(header);

  const body = document.createElement('div');
  body.className = 'mail-view-body';
  body.textContent = mail.body?.content || mail.body || '';
  view.appendChild(body);

  visualContentArea.appendChild(view);
}

function renderMailComposeForm(messages) {
  visualContentArea.innerHTML = '';
  const form = document.createElement('div');
  form.className = 'mail-view';
  form.innerHTML = `
    <button class="mail-view-back" id="compose-back-btn">&larr; Cancel</button>
    <h2 class="mail-view-title" style="margin-bottom: 12px; padding-bottom: 6px;">New Message</h2>
    <div style="display: flex; flex-direction: column; gap: 10px; flex: 1;">
      <div class="settings-group">
        <label>To:</label>
        <input type="text" id="compose-to" placeholder="recipient@contoso.com" style="background: var(--bg-primary); border: var(--border-dim); color: var(--text-main); padding: 6px 10px; border-radius: 4px; outline: none;">
      </div>
      <div class="settings-group">
        <label>Subject:</label>
        <input type="text" id="compose-subject" placeholder="Enter subject line..." style="background: var(--bg-primary); border: var(--border-dim); color: var(--text-main); padding: 6px 10px; border-radius: 4px; outline: none;">
      </div>
      <div class="settings-group" style="flex: 1; display: flex; flex-direction: column;">
        <label>Body:</label>
        <textarea id="compose-body" placeholder="Type your email contents here..." style="background: var(--bg-primary); border: var(--border-dim); color: var(--text-main); padding: 10px; border-radius: 4px; outline: none; flex: 1; resize: none; font-family: inherit; font-size: 0.85rem;"></textarea>
      </div>
      <button class="btn primary-btn" id="compose-send-btn" style="align-self: flex-end; margin-top: 8px;">Send Command</button>
    </div>
  `;

  form.querySelector('#compose-back-btn').addEventListener('click', () => {
    visualContentArea.innerHTML = '';
    renderMailApp(messages);
  });

  form.querySelector('#compose-send-btn').addEventListener('click', () => {
    const to = form.querySelector('#compose-to').value.trim();
    const sub = form.querySelector('#compose-subject').value.trim();
    const body = form.querySelector('#compose-body').value.trim();

    if (!to || !body) {
      alert("Please fill in To and Body fields.");
      return;
    }

    if (terminalConsole) {
      const input = document.getElementById('terminal-input');
      input.value = `mail --send --to "${to}" --subject "${sub}" --body "${body}"`;
      input.focus();
    }
    
    // Return to mail list
    visualContentArea.innerHTML = '';
    renderMailApp(messages);
  });

  visualContentArea.appendChild(form);
}

function renderCalendarApp(events) {
  const container = document.createElement('div');
  container.className = 'calendar-app';

  const toolbar = document.createElement('div');
  toolbar.className = 'calendar-header';
  toolbar.innerHTML = `
    <h3 style="font-size: 1.1rem;">Calendar Schedule</h3>
    <button class="btn-compose" id="btn-add-event">
      <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
      Add Event
    </button>
  `;

  // Draw a mini calendar grid (representing current week)
  const grid = document.createElement('div');
  grid.className = 'calendar-grid';
  
  // Headers
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  days.forEach(d => {
    const h = document.createElement('div');
    h.className = 'calendar-day-header';
    h.textContent = d;
    grid.appendChild(h);
  });

  // Calculate days of current week
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay());

  for (let i = 0; i < 7; i++) {
    const cellDate = new Date(startOfWeek);
    cellDate.setDate(startOfWeek.getDate() + i);

    const cell = document.createElement('div');
    cell.className = 'calendar-cell';
    if (cellDate.toDateString() === today.toDateString()) {
      cell.classList.add('today');
    }

    // Check if this day has events
    const hasEvents = events.some(ev => {
      const evDate = new Date(ev.start?.dateTime || ev.start);
      return evDate.toDateString() === cellDate.toDateString();
    });

    cell.innerHTML = `<span class="calendar-date-num">${cellDate.getDate()}</span>`;
    if (hasEvents) {
      cell.innerHTML += `<span class="calendar-indicator-dot"></span>`;
    }

    grid.appendChild(cell);
  }

  // Events list
  const list = document.createElement('div');
  list.className = 'calendar-list';

  if (events.length === 0) {
    list.innerHTML = `<div class="empty-state"><p>No upcoming events scheduled.</p></div>`;
  } else {
    // Sort oldest to newest
    const sorted = [...events].sort((a,b) => new Date(a.start?.dateTime || a.start) - new Date(b.start?.dateTime || b.start));
    sorted.forEach(ev => {
      const item = document.createElement('div');
      item.className = 'event-item';
      
      const startTime = new Date(ev.start?.dateTime || ev.start).toLocaleString([], {month: 'short', day: 'numeric', hour: '2-digit', minute:'2-digit'});
      
      item.innerHTML = `
        <div class="event-details">
          <span class="event-subject">${ev.subject}</span>
          <span class="event-time">${startTime} @ ${ev.location?.displayName || 'Online'}</span>
        </div>
        <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color: var(--accent-cyan);"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
      `;
      list.appendChild(item);
    });
  }

  toolbar.querySelector('#btn-add-event').addEventListener('click', () => renderCalendarForm(events));

  container.appendChild(toolbar);
  container.appendChild(grid);
  container.appendChild(list);
  visualContentArea.appendChild(container);
}

function renderCalendarForm(events) {
  visualContentArea.innerHTML = '';
  const form = document.createElement('div');
  form.className = 'mail-view';
  form.innerHTML = `
    <button class="mail-view-back" id="cal-back-btn">&larr; Cancel</button>
    <h2 class="mail-view-title">New Event</h2>
    <div style="display: flex; flex-direction: column; gap: 12px;">
      <div class="settings-group">
        <label>Subject/Title:</label>
        <input type="text" id="event-subject-input" placeholder="Project sync, 1:1, etc." style="background: var(--bg-primary); border: var(--border-dim); color: var(--text-main); padding: 6px 10px; border-radius: 4px; outline: none;">
      </div>
      <div class="settings-group">
        <label>Start Date/Time:</label>
        <input type="datetime-local" id="event-start-input" style="background: var(--bg-primary); border: var(--border-dim); color: var(--text-main); padding: 6px 10px; border-radius: 4px; outline: none;">
      </div>
      <div class="settings-group">
        <label>End Date/Time:</label>
        <input type="datetime-local" id="event-end-input" style="background: var(--bg-primary); border: var(--border-dim); color: var(--text-main); padding: 6px 10px; border-radius: 4px; outline: none;">
      </div>
      <div class="settings-group">
        <label>Location:</label>
        <input type="text" id="event-location-input" placeholder="Teams Meeting, Room 10" style="background: var(--bg-primary); border: var(--border-dim); color: var(--text-main); padding: 6px 10px; border-radius: 4px; outline: none;">
      </div>
      <button class="btn primary-btn" id="event-submit-btn" style="align-self: flex-end; margin-top: 10px;">Schedule Event</button>
    </div>
  `;

  // Pre-fill datetime-local inputs
  const start = new Date();
  start.setHours(start.getHours() + 1, 0, 0, 0);
  const end = new Date(start);
  end.setMinutes(end.getMinutes() + 30);
  
  form.querySelector('#event-start-input').value = start.toISOString().substring(0, 16);
  form.querySelector('#event-end-input').value = end.toISOString().substring(0, 16);

  form.querySelector('#cal-back-btn').addEventListener('click', () => {
    visualContentArea.innerHTML = '';
    renderCalendarApp(events);
  });

  form.querySelector('#event-submit-btn').addEventListener('click', () => {
    const subject = form.querySelector('#event-subject-input').value.trim();
    const startVal = form.querySelector('#event-start-input').value;
    const endVal = form.querySelector('#event-end-input').value;
    const loc = form.querySelector('#event-location-input').value.trim() || 'Teams Meeting';

    if (!subject || !startVal || !endVal) {
      alert("Please fill in subject, start, and end times.");
      return;
    }

    if (terminalConsole) {
      const input = document.getElementById('terminal-input');
      input.value = `calendar --add --subject "${subject}" --start "${startVal}" --end "${endVal}" --location "${loc}"`;
      input.focus();
    }
    
    visualContentArea.innerHTML = '';
    renderCalendarApp(events);
  });

  visualContentArea.appendChild(form);
}

function renderDriveApp(items) {
  const container = document.createElement('div');
  container.className = 'drive-app';
  
  const header = document.createElement('div');
  header.style.display = 'flex';
  header.style.justifyContent = 'space-between';
  header.style.alignItems = 'center';
  header.innerHTML = `
    <div class="drive-breadcrumbs">drive:/root</div>
    <div style="display: flex; gap: 8px;">
      <button class="btn-compose" id="btn-mkdir" style="border-color: rgba(255, 179, 0, 0.3); color: var(--accent-amber);">+ New Folder</button>
      <button class="btn-compose" id="btn-upload">+ Upload File</button>
    </div>
  `;

  const list = document.createElement('div');
  list.className = 'drive-list';

  if (items.length === 0) {
    list.innerHTML = `<div class="empty-state"><p>This directory is empty.</p></div>`;
  } else {
    items.forEach(it => {
      const row = document.createElement('div');
      row.className = 'drive-item';
      
      const icon = it.folder || it.type === 'folder' 
        ? `<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" class="drive-icon folder"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>`
        : `<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="drive-icon"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>`;

      row.innerHTML = `
        <div class="drive-item-meta">
          ${icon}
          <span class="drive-name">${it.name}</span>
        </div>
        <span class="drive-size">${it.size || '--'}</span>
      `;
      list.appendChild(row);
    });
  }

  // Create folder prompt
  header.querySelector('#btn-mkdir').addEventListener('click', () => {
    const name = prompt("Enter new folder name:");
    if (!name) return;
    if (terminalConsole) {
      const input = document.getElementById('terminal-input');
      input.value = `drive --mkdir --name "${name}"`;
      input.focus();
    }
  });

  // Create upload file prompt
  header.querySelector('#btn-upload').addEventListener('click', () => {
    const name = prompt("Enter mock file name to upload:");
    if (!name) return;
    if (terminalConsole) {
      const input = document.getElementById('terminal-input');
      input.value = `drive --upload --name "${name}"`;
      input.focus();
    }
  });

  container.appendChild(header);
  container.appendChild(list);
  visualContentArea.appendChild(container);
}

function formatMarkdownToHTML(md) {
  let lines = md.split('\n');
  let html = '';
  let inTable = false;
  let tableHeaderParsed = false;
  
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    
    // Check if it's a table row
    if (line.startsWith('|') && line.endsWith('|')) {
      if (!inTable) {
        inTable = true;
        tableHeaderParsed = false;
        html += '<table style="width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 0.8rem; border: var(--border-dim); box-shadow: var(--shadow-neon-dim);">';
      }
      
      // Split by '|' and clean columns
      let cols = line.split('|').map(c => c.trim()).filter((c, idx, arr) => idx > 0 && idx < arr.length - 1);
      
      // Skip alignment rows e.g. |---|---|
      if (line.replace(/[\s|:-]/g, '').length === 0) {
        continue;
      }
      
      if (!tableHeaderParsed) {
        tableHeaderParsed = true;
        html += '<thead><tr style="background: rgba(0, 240, 255, 0.06); border-bottom: 1px solid rgba(0, 240, 255, 0.2);">';
        cols.forEach(col => {
          html += `<th style="padding: 10px 12px; text-align: left; font-weight: 600; color: var(--accent-cyan); border: var(--border-dim);">${col}</th>`;
        });
        html += '</tr></thead><tbody>';
      } else {
        html += '<tr style="border-bottom: var(--border-dim);">';
        cols.forEach(col => {
          html += `<td style="padding: 10px 12px; color: var(--text-main); border: var(--border-dim);">${col}</td>`;
        });
        html += '</tr>';
      }
    } else {
      if (inTable) {
        inTable = false;
        html += '</tbody></table>';
      }
      
      // Parse list items
      if (line.startsWith('* ') || line.startsWith('- ')) {
        const content = line.substring(2);
        html += `<li style="margin-left: 14px; margin-bottom: 4px; list-style-type: square; color: var(--text-main);">${content}</li>`;
      } else if (line.startsWith('### ')) {
        html += `<h4 style="color: var(--accent-cyan); font-weight: 600; margin-top: 14px; margin-bottom: 6px;">${line.substring(4)}</h4>`;
      } else if (line.startsWith('## ')) {
        html += `<h4 style="color: var(--accent-cyan); font-weight: 600; margin-top: 14px; margin-bottom: 6px;">${line.substring(3)}</h4>`;
      } else if (line) {
        html += `<p style="margin-bottom: 8px;">${line}</p>`;
      } else {
        html += '<br>';
      }
    }
  }
  
  if (inTable) {
    html += '</tbody></table>';
  }
  
  // Format bold text
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  // Format inline code backticks
  html = html.replace(/`(.*?)`/g, '<code class="code-font" style="background: rgba(255, 255, 255, 0.06); padding: 2px 5px; border-radius: 4px; font-size: 0.78rem; color: var(--accent-cyan); border: 1px solid rgba(0, 240, 255, 0.1);">$1</code>');
  
  return html;
}

function renderErrorCard(errorData) {
  const card = document.createElement('div');
  card.className = 'error-resolver-view';
  card.style.display = 'flex';
  card.style.flexDirection = 'column';
  card.style.gap = '16px';
  card.style.width = '100%';
  card.style.padding = '20px';

  card.innerHTML = `
    <div style="display: flex; align-items: center; gap: 10px; color: var(--accent-red); margin-bottom: 8px;">
      <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="box-shadow: 0 0 10px rgba(255, 51, 102, 0.2); border-radius: 50%; padding: 4px;"><polygon points="7.86 2 16.14 2 22 7.86 22 16.14 16.14 22 7.86 22 2 16.14 2 7.86 7.86 2"></polygon><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>
      <h2 style="font-size: 1.3rem; font-weight: 700; text-shadow: 0 0 10px rgba(255,51,102,0.3); margin: 0;">API Error Resolver</h2>
    </div>
    
    <div class="api-logger-card" style="border-color: rgba(255, 51, 102, 0.2); background: rgba(255, 51, 102, 0.02); margin: 0;">
      <div class="api-logger-body" style="padding: 12px 14px;">
        <div class="log-row"><span class="label" style="color: var(--text-dim); margin-right: 8px;">Request:</span> <span class="value code-font" style="color: var(--text-main); font-weight: 600;">${errorData.method} ${errorData.endpoint}</span></div>
        <div class="log-row" style="margin-top: 6px;"><span class="label" style="color: var(--text-dim); margin-right: 8px;">Error:</span> <span class="value code-font" style="color: var(--accent-red); font-size: 0.8rem; word-break: break-all;">${errorData.message}</span></div>
      </div>
    </div>

    <div class="api-logger-card" style="border-color: rgba(0, 240, 255, 0.2); background: rgba(0, 240, 255, 0.02); margin: 0;">
      <div class="api-logger-header" style="border-bottom: 1px solid rgba(0, 240, 255, 0.1); padding: 8px 12px; display: flex; align-items: center; justify-content: space-between;">
        <span class="api-logger-badge" style="background: rgba(0, 240, 255, 0.1); color: var(--accent-cyan);">AZURE AI SOLUTIONS</span>
        <span class="api-logger-status success" id="ai-analysis-loader" style="font-size: 0.75rem; text-transform: uppercase;">ANALYZING...</span>
      </div>
      <div class="api-logger-body" id="ai-error-analysis-body" style="padding: 14px; color: var(--text-main); font-size: 0.85rem; line-height: 1.5; white-space: pre-wrap; font-family: var(--font-sans);">
        <span class="pulse-icon">Consulting Azure OpenAI model for solution steps...</span>
      </div>
    </div>
  `;
  
  visualContentArea.appendChild(card);

  // Switch to the Visual Explorer tab automatically
  const tabVisual = document.getElementById('tab-visual');
  if (tabVisual) tabVisual.click();

  // Trigger Azure OpenAI analysis asynchronously
  import('./aiService.js').then(module => {
    const accounts = msalInstance ? msalInstance.getAllAccounts() : [];
    const tenantId = accounts[0]?.homeAccountId?.split('.')[1] || '';
    const accountType = tenantId === '9188040d-6c67-4c5b-b112-36a304b66dad' ? 'Personal Microsoft Account (MSA)' : `Work/School Directory Tenant (${tenantId})`;

    module.explainApiError(errorData.endpoint, errorData.message, accountType).then(analysis => {
      const loader = document.getElementById('ai-analysis-loader');
      if (loader) {
        loader.textContent = 'ANALYSIS COMPLETE';
        loader.className = 'api-logger-status success';
      }
      const body = document.getElementById('ai-error-analysis-body');
      if (body) {
        body.innerHTML = formatMarkdownToHTML(analysis);
      }
    });
  }).catch(err => {
    console.error("Failed to load aiService:", err);
    const body = document.getElementById('ai-error-analysis-body');
    if (body) body.textContent = "Could not load AI module to generate solutions.";
  });
}
