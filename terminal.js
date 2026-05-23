// Command Argument Parser supporting quotes
export function parseArgs(rawInput) {
  const args = [];
  let current = "";
  let insideQuotes = false;
  let quoteChar = "";

  for (let i = 0; i < rawInput.length; i++) {
    const char = rawInput[i];
    if ((char === '"' || char === "'") && (i === 0 || rawInput[i - 1] !== '\\')) {
      if (insideQuotes && char === quoteChar) {
        insideQuotes = false;
      } else if (!insideQuotes) {
        insideQuotes = true;
        quoteChar = char;
      }
    } else if (char === ' ' && !insideQuotes) {
      if (current) {
        args.push(current);
        current = "";
      }
    } else {
      current += char;
    }
  }
  if (current) {
    args.push(current);
  }

  const cmd = args.shift() || "";
  const flags = {};
  const params = [];

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg.startsWith('--')) {
      const key = arg.slice(2);
      if (args[i + 1] && !args[i + 1].startsWith('-')) {
        flags[key] = args[i + 1];
        i++;
      } else {
        flags[key] = true;
      }
    } else if (arg.startsWith('-')) {
      const key = arg.slice(1);
      if (args[i + 1] && !args[i + 1].startsWith('-')) {
        flags[key] = args[i + 1];
        i++;
      } else {
        flags[key] = true;
      }
    } else {
      params.push(arg);
    }
  }

  return { cmd, flags, params };
}

export class TerminalConsole {
  constructor(inputElement, logContainer, commandRegistry) {
    this.input = inputElement;
    this.log = logContainer;
    this.commands = commandRegistry;
    
    this.history = JSON.parse(localStorage.getItem('graphshell_history') || '[]');
    this.historyIndex = this.history.length;
    this.maxHistory = 100;
    
    this.autocompleteOptions = [
      'help', 'login', 'logout', 'me', 'users', 'groups', 'mail', 'calendar', 'drive', 'request', 'ai', 'settings', 'config', 'clear'
    ];

    this.initEvents();
  }

  initEvents() {
    this.input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        this.executeCommand();
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        this.navigateHistory(-1);
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        this.navigateHistory(1);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        this.triggerAutocomplete();
      }
    });

    // Keep focus in terminal if user clicks inside it
    document.getElementById('terminal-section').addEventListener('click', () => {
      this.input.focus();
    });
  }

  writeLine(text, type = 'default') {
    const line = document.createElement('div');
    line.className = `log-line ${type}`;
    line.textContent = text;
    
    const inputContainer = this.log.querySelector('.terminal-input-container');
    if (inputContainer) {
      this.log.insertBefore(line, inputContainer);
    } else {
      this.log.appendChild(line);
    }
    
    this.log.scrollTop = this.log.scrollHeight;
  }

  clear() {
    const inputContainer = this.log.querySelector('.terminal-input-container');
    this.log.innerHTML = '';
    if (inputContainer) {
      this.log.appendChild(inputContainer);
    }
    this.input.focus();
  }

  handleCommand(rawInput) {
    if (!rawInput) return;
    this.input.value = rawInput;
    this.executeCommand();
  }

  executeCommand() {
    const rawInput = this.input.value.trim();
    this.input.value = '';

    if (!rawInput) return;

    // Log the input
    this.writeLine(`graphshell$ ${rawInput}`, 'cmd-input');

    // Add to history
    this.history.push(rawInput);
    if (this.history.length > this.maxHistory) this.history.shift();
    this.historyIndex = this.history.length;
    localStorage.setItem('graphshell_history', JSON.stringify(this.history));

    // Parse and run
    const parsed = parseArgs(rawInput);
    const cmdName = parsed.cmd.toLowerCase();

    if (cmdName === 'clear') {
      this.clear();
      return;
    }

    if (this.commands[cmdName]) {
      try {
        this.commands[cmdName](parsed.flags, parsed.params, this);
      } catch (err) {
        this.writeLine(`Error: ${err.message}`, 'error');
      }
    } else {
      this.writeLine(`Command not found: '${parsed.cmd}'. Type 'help' for a list of commands.`, 'error');
    }
  }

  navigateHistory(direction) {
    if (this.history.length === 0) return;

    this.historyIndex += direction;
    if (this.historyIndex < 0) {
      this.historyIndex = 0;
    } else if (this.historyIndex > this.history.length) {
      this.historyIndex = this.history.length;
    }

    if (this.historyIndex === this.history.length) {
      this.input.value = '';
    } else {
      this.input.value = this.history[this.historyIndex];
    }
  }

  triggerAutocomplete() {
    const currentVal = this.input.value.trim();
    if (!currentVal) return;

    const matches = this.autocompleteOptions.filter(cmd => cmd.startsWith(currentVal.toLowerCase()));
    
    if (matches.length === 1) {
      this.input.value = matches[0] + ' ';
    } else if (matches.length > 1) {
      this.writeLine(`Possible matches: ${matches.join(', ')}`, 'info');
    }
  }
}
