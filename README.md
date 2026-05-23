# 💻 GraphShell: MS Graph Terminal Portal & AI Copilot

An interactive, high-fidelity terminal portal and AI Copilot for Microsoft Graph API and Azure AD directory services. GraphShell provides a real-time command-line interface directly in your browser, combined with rich visual widgets, an active API transaction logger, and AI-powered natural language query translation.

---

## 🏛️ Core Architecture

The application is structured as a decoupled, responsive client-side Single Page Application (SPA) driven by standard ES Modules.

```mermaid
graph TD
    User([User Client]) <--> UI[GraphShell Dashboard UI]
    UI <--> Terminal[Terminal CLI terminal.js]
    UI <--> Tabs[Visual Tabs & AI Chat Panel]
    Terminal <--> Cmd[Command Router commands.js]
    Cmd <--> MSAL[MSAL Browser Auth]
    Cmd <--> GraphAPI[Microsoft Graph API REST Endpoint]
    Tabs <--> AIService[AI Copilot Service aiService.js]
    AIService <--> AzureOpenAI[Azure OpenAI Service API]
    GraphAPI -.-> APILog[API Logger Widget]
    GraphAPI -.-> ErrorResolver[AI Error Diagnostics Resolver]
```

---

## 🔄 Sequence Flow Diagram

The sequence below illustrates how user commands are captured, authenticated, resolved against Microsoft Graph, and how the AI Copilot translates prompts:

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant UI as GraphShell UI
    participant Term as Terminal Console
    participant Cmd as Command Router
    participant MSAL as MSAL.js
    participant Graph as Microsoft Graph API
    participant AI as Azure OpenAI Copilot

    User->>UI: Types command (e.g., "users --list")
    UI->>Term: Process input
    Term->>Cmd: Execute "users --list"
    alt Not Logged In
        Cmd->>MSAL: Acquire Token / Login Redirect
        MSAL-->>Cmd: Access Token
    end
    Cmd->>Graph: HTTP GET /v1.0/users (Bearer Token)
    alt Success
        Graph-->>Cmd: JSON Response Data
        Cmd->>UI: Update Visual Explorer & JSON tab
    else Failure (e.g. MSA Account error)
        Graph-->>Cmd: HTTP Error (400 Bad Request)
        Cmd->>AI: Fetch Error explanation & solution
        AI-->>Cmd: Resolution details
        Cmd->>UI: Render Error Resolver Panel
    end
```

---

## 📂 Project Directory Structure

Below is the layout of the GraphShell codebase:

```
graph-shell/
├── Footer/
│   ├── Footer.css               # CSS styles for the footer component
│   └── Footer.jsx               # React wrapper of the footer component
├── aiService.js                 # Azure OpenAI interface & natural language translation engine
├── app.js                       # Main application lifecycle, state controller, & UI events
├── azure_openai_setup.md        # CLI script template to deploy Azure OpenAI instances
├── commands.js                  # Microsoft Graph command definitions and request router
├── index.css                    # Main application stylesheet (CRT scanlines, glassmorphism, responsive grid)
├── index.html                   # HTML5 page layout and layout panels
├── logo.png                     # Glowing tech dashboard logo & favicon source
├── manifest.json                # PWA progressive web app configuration
├── sw.js                        # Service worker caching engine
└── terminal.js                  # Terminal input line controller and boot emulator
```

---

## 🛠️ Technology Stack & Versions

GraphShell is optimized for speed, security, and visuals by running fully client-side:

| Technology | Layer | Version / Source | Description |
| :--- | :--- | :--- | :--- |
| ![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=flat-square&logo=html5&logoColor=white) | Core Markup | Living Standard | Semantic document outline with a responsive dual-pane flex layout. |
| ![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=flat-square&logo=css3&logoColor=white) | Styling System | CSS3 / Custom Variables | Custom CRT scanner styling, glassmorphism overlays, and dark-mode gradients. |
| ![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat-square&logo=javascript&logoColor=black) | Logic / Controller | ES6+ Modules | Event handling, state management, and async fetch operations. |
| ![Microsoft](https://img.shields.io/badge/Microsoft_Entra-0078D4?style=flat-square&logo=microsoft&logoColor=white)| Authentication | `v2.35.0` (CDN) | Microsoft Authentication Library for secure Entra ID OAuth sign-ins. |
| ![Azure](https://img.shields.io/badge/Azure_OpenAI-0089D6?style=flat-square&logo=microsoftazure&logoColor=white)| AI Engine | `gpt-4.1-mini` (2025-04-14) | Cognitive model deployment for natural language CLI routing. |
| ![Google](https://img.shields.io/badge/Google_Fonts-4285F4?style=flat-square&logo=google&logoColor=white)| Typography | `Inter` & `Fira Code` | Modern UI typography and fixed-width command typography. |

---

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

---

## Author

### 👤 Francis Ponnu Cruz I
> **Azure Cloud & DevOps Engineer | Microsoft Certified Trainer (MCT)**

#### 🌐 Connect with Me:
[![GitHub](https://img.shields.io/badge/GitHub-ajf013-181717?style=flat-square&logo=github)](https://github.com/ajf013)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Francis_Cruz-0A66C2?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/ajf013-francis-cruz/)
[![Twitter/X](https://img.shields.io/badge/X-@Itsme__Ajf013-000000?style=flat-square&logo=x)](https://x.com/Itsme_Ajf013)
[![Website](https://img.shields.io/badge/Website-fcruz.org-2D3748?style=flat-square&logo=googlechrome&logoColor=white)](https://fcruz.org)
[![Linktree](https://img.shields.io/badge/Linktree-AJF013-39E09B?style=flat-square&logo=linktree&logoColor=white)](https://linktr.ee/AJF013)
