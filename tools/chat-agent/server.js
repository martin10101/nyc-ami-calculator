'use strict';

const express = require('express');
const OpenAI = require('openai');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const AGENT_ROOT = process.env.AGENT_ROOT || 'C:\\AMI_Optix_Agent';
const REPO_ROOT  = process.env.REPO_ROOT  || '';
const PORT       = parseInt(process.env.PORT || '3000', 10);

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY || '' });

// ── Conversation history (persisted to disk) ──────────────────────────────────
const HISTORY_FILE = path.join(AGENT_ROOT, 'state', 'chat-history.json');

function loadHistory() {
  try {
    if (fs.existsSync(HISTORY_FILE)) return JSON.parse(fs.readFileSync(HISTORY_FILE, 'utf8'));
  } catch {}
  return [];
}

function saveHistory(history) {
  try {
    const dir = path.dirname(HISTORY_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(HISTORY_FILE, JSON.stringify(history.slice(-200), null, 2), 'utf8');
  } catch {}
}

app.get('/api/history', (req, res) => res.json(loadHistory()));

// ── Session state ─────────────────────────────────────────────────────────────
const sessions = new Map();

// ── Tool definitions ──────────────────────────────────────────────────────────
const TOOLS = [
  {
    type: 'function',
    function: {
      name: 'run_acceptance_tests',
      description: 'Run all acceptance test scenarios and return the full results.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'run_full_refresh',
      description: 'Full cycle: bootstrap workspace, rebuild the add-in, then run acceptance tests.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'list_source_files',
      description: 'List all editable VBA and PowerShell source files in the repository.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_source_file',
      description: 'Read the full contents of a source file.',
      parameters: {
        type: 'object',
        properties: {
          relative_path: { type: 'string', description: 'Path relative to repo root' }
        },
        required: ['relative_path']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'edit_source_file',
      description: 'Make a targeted edit to a source file by replacing an exact string.',
      parameters: {
        type: 'object',
        properties: {
          relative_path: { type: 'string' },
          old_string:    { type: 'string' },
          new_string:    { type: 'string' }
        },
        required: ['relative_path', 'old_string', 'new_string']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_results',
      description: 'Read the latest acceptance test results JSON.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_build_result',
      description: 'Read the latest build result JSON.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'compile_check',
      description: 'After editing source files, run this before full tests. It rebuilds the add-in and calls a no-op function to verify the project compiles cleanly. Returns COMPILE_OK or COMPILE_ERROR with the exact error text. Much faster than a full refresh.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'ask_user',
      description: 'Pause and ask the user a question when you cannot proceed without their input.',
      parameters: {
        type: 'object',
        properties: {
          question: { type: 'string' }
        },
        required: ['question']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'flag_server_change',
      description: 'Flag that a server-side file (app.py, solver.py, etc.) needs updating. This machine cannot push to GitHub. The user will coordinate the push from their dev machine. Use this whenever you identify a bug or needed change in server-side Python code.',
      parameters: {
        type: 'object',
        properties: {
          file_path:    { type: 'string', description: 'Server-side file that needs changing (e.g. app.py, ami_optix/solver.py)' },
          description:  { type: 'string', description: 'What needs to change and why' },
          severity:     { type: 'string', enum: ['blocker', 'important', 'nice-to-have'], description: 'How urgent is this change' }
        },
        required: ['file_path', 'description', 'severity']
      }
    }
  }
];

// ── PowerShell runner with parallel dialog watcher ────────────────────────────
function runPowershell(scriptPath, args, sessionId, timeoutMs, send) {
  if (!timeoutMs) timeoutMs = 60000;
  return new Promise((resolve) => {
    const session = sessions.get(sessionId);

    // Main process
    const ps = spawn('powershell', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, ...args
    ]);
    if (session) session.activeProcess = ps;

    let stdout = '';
    let stderr = '';
    let finished = false;

    // Dialog watcher runs in parallel
    const watcherScript = path.join(__dirname, 'Watch-ExcelDialog.ps1');
    let watcher = null;
    if (fs.existsSync(watcherScript)) {
      watcher = spawn('powershell', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', watcherScript,
        '-TimeoutMs', String(timeoutMs), '-PollMs', '300'
      ]);
      watcher.stdout.on('data', d => {
        const line = d.toString().trim();
        if (line.startsWith('DIALOG_DISMISSED:')) {
          const errorText = line.slice('DIALOG_DISMISSED:'.length).trim();
          if (send) send('dialog_dismissed', errorText);
          stdout += `\n[Excel dialog auto-dismissed. Error text: ${errorText}]`;
        }
      });
    }

    const finish = (code) => {
      if (finished) return;
      finished = true;
      if (session) session.activeProcess = null;
      if (watcher) { try { watcher.kill(); } catch {} }
      resolve({ code, stdout, stderr });
    };

    ps.stdout.on('data', d => { stdout += d.toString(); });
    ps.stderr.on('data', d => { stderr += d.toString(); });
    ps.on('close', finish);

    const timer = setTimeout(() => {
      if (!finished) {
        try { ps.kill(); } catch {}
        resolve({
          code: -1,
          stdout,
          stderr: (stderr || '') + '\n[TIMED OUT — Excel may have an open dialog. Please close Excel and try again.]'
        });
      }
    }, timeoutMs);

    ps.on('close', () => clearTimeout(timer));
  });
}

// ── Kill lingering Excel processes ────────────────────────────────────────────
function killExcel() {
  return new Promise(resolve => {
    const ps = spawn('powershell', ['-NoProfile', '-Command',
      'Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Milliseconds 800']);
    ps.on('close', resolve);
    setTimeout(() => { try { ps.kill(); } catch {} resolve(); }, 5000);
  });
}

// ── File helpers ──────────────────────────────────────────────────────────────
function walkDir(dir, repoRoot, results) {
  if (!fs.existsSync(dir)) return;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      walkDir(full, repoRoot, results);
    } else if (/\.(bas|cls|frm|ps1|psm1|json)$/i.test(entry.name)) {
      results.push(path.relative(repoRoot, full).replace(/\\/g, '/'));
    }
  }
}

// ── Tool executor ─────────────────────────────────────────────────────────────
async function executeTool(name, args, sessionId, send) {
  const session = sessions.get(sessionId);

  if (name === 'ask_user') {
    return new Promise((resolve, reject) => {
      if (!session || session.aborted) { reject(new Error('Session aborted')); return; }
      session.resolvePause = resolve;
      session.rejectPause  = reject;
      send('pause', { question: args.question });
    });
  }

  switch (name) {
    case 'run_acceptance_tests': {
      const script = path.join(AGENT_ROOT, 'scripts', 'Invoke-AmiOptixAcceptance.ps1');
      if (!fs.existsSync(script)) return 'Acceptance script not found. Run a full refresh first.';
      await killExcel();
      const r = await runPowershell(script, ['-AgentRoot', AGENT_ROOT], sessionId, 300000, send);
      return (r.stdout + (r.stderr ? '\nSTDERR: ' + r.stderr : '')).trim();
    }

    case 'compile_check': {
      const script = path.join(AGENT_ROOT, 'scripts', 'Invoke-AmiOptixCompileCheck.ps1');
      if (!fs.existsSync(script)) return 'Compile check script not found. Run a full refresh first to bootstrap the workspace.';
      await killExcel();
      const r = await runPowershell(script, ['-AgentRoot', AGENT_ROOT], sessionId, 120000, send);
      return (r.stdout + (r.stderr ? '\nSTDERR: ' + r.stderr : '')).trim();
    }

    case 'run_full_refresh': {
      const repoScript = REPO_ROOT
        ? path.join(REPO_ROOT, 'tools', 'excel-agent', 'Refresh-AmiOptixAgent.ps1')
        : path.join(AGENT_ROOT, 'scripts', 'Refresh-AmiOptixAgent.ps1');
      const psArgs = ['-AgentRoot', AGENT_ROOT];
      if (REPO_ROOT) psArgs.push('-RepoRoot', REPO_ROOT);
      await killExcel();
      const r = await runPowershell(repoScript, psArgs, sessionId, 600000, send);
      return (r.stdout + (r.stderr ? '\nSTDERR: ' + r.stderr : '')).trim();
    }

    case 'list_source_files': {
      if (!REPO_ROOT) return 'REPO_ROOT is not configured.';
      const files = [];
      walkDir(path.join(REPO_ROOT, 'excel-addin', 'src'), REPO_ROOT, files);
      walkDir(path.join(REPO_ROOT, 'tools', 'excel-agent'), REPO_ROOT, files);
      return files.length ? files.join('\n') : 'No source files found.';
    }

    case 'read_source_file': {
      if (!REPO_ROOT) return 'REPO_ROOT is not configured.';
      const filePath = path.join(REPO_ROOT, args.relative_path.replace(/\//g, path.sep));
      try { return fs.readFileSync(filePath, 'utf8'); }
      catch (e) { return 'Error reading file: ' + e.message; }
    }

    case 'edit_source_file': {
      if (!REPO_ROOT) return 'REPO_ROOT is not configured.';
      const filePath = path.join(REPO_ROOT, args.relative_path.replace(/\//g, path.sep));
      try {
        const content = fs.readFileSync(filePath, 'utf8');
        if (!content.includes(args.old_string)) {
          return 'ERROR: old_string not found in file — no changes made.';
        }
        const updated = content.replace(args.old_string, args.new_string);
        fs.writeFileSync(filePath, updated, 'utf8');
        const preview = args.old_string.slice(0, 120).replace(/\n/g, '↵');
        send('file_edited', { path: args.relative_path, preview });
        return 'OK: edited ' + args.relative_path;
      } catch (e) {
        return 'Error editing file: ' + e.message;
      }
    }

    case 'read_results': {
      const p = path.join(AGENT_ROOT, 'artifacts', 'acceptance-result.json');
      try { return fs.readFileSync(p, 'utf8'); }
      catch (e) { return 'No acceptance results found: ' + e.message; }
    }

    case 'read_build_result': {
      const p = path.join(AGENT_ROOT, 'artifacts', 'build-result.json');
      try { return fs.readFileSync(p, 'utf8'); }
      catch (e) { return 'No build results found: ' + e.message; }
    }

    case 'flag_server_change': {
      const entry = {
        timestamp: new Date().toISOString(),
        file: args.file_path,
        description: args.description,
        severity: args.severity || 'important'
      };
      const flagFile = path.join(AGENT_ROOT, 'state', 'pending-server-changes.json');
      let existing = [];
      try { existing = JSON.parse(fs.readFileSync(flagFile, 'utf8')); } catch {}
      existing.push(entry);
      try {
        const dir = path.dirname(flagFile);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(flagFile, JSON.stringify(existing, null, 2), 'utf8');
      } catch {}
      send('server_change_flagged', entry);
      return `Flagged: ${args.file_path} (${args.severity}) — "${args.description}". The user has been notified and will push this change via GitHub.`;
    }

    default:
      return 'Unknown tool: ' + name;
  }
}

// ── Load full project context ────────────────────────────────────────────────
let AGENT_CONTEXT_TEXT = '';
try {
  const ctxPath = path.join(__dirname, 'AGENT_CONTEXT.md');
  if (fs.existsSync(ctxPath)) {
    AGENT_CONTEXT_TEXT = fs.readFileSync(ctxPath, 'utf8');
  }
} catch {}

// ── System prompt ─────────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are an autonomous repair and development agent for AMI Optix — a NYC affordable housing AMI calculator built as a VBA Excel add-in with a PowerShell automation layer.

NARRATE every step before you take it: one sentence saying what you will do and why. After each result, one sentence on what you found.

══════════════════════════════════════════
GOLDEN RULE: Never run acceptance tests while there is a known compile error.
Fix compile → COMPILE_OK → then test.
══════════════════════════════════════════

LOOP FOR EVERY TASK:
1. Run compile_check first (fast, ~30s). If COMPILE_ERROR → jump to step 3.
2. Run acceptance tests. If all pass → summarize and stop.
3. Read the full error message. Find the exact file and line. Fix it with edit_source_file.
4. Run compile_check. If still COMPILE_ERROR → back to step 3 with the new error.
5. When compile_check returns COMPILE_OK → back to step 2.
6. Same error appears 3 times in a row without any change → call ask_user.

COMPILE ERRORS — how to handle:
- "Variable not defined: X" → find where X is used before its Dim declaration. Either move the Dim earlier, or remove the premature use. Read the file first to see the actual code.
- "Sub or Function not defined: X" → X is called but doesn't exist. Add the function, or remove the call.
- "Type mismatch" or "Object required" → wrong type assigned; check Set vs plain assignment for objects.
- When test output contains "[Excel dialog auto-dismissed. Error text: X]" → X is the real error. Fix the code that caused X, then run compile_check.
- After any edit, ALWAYS run compile_check before running tests.

EXCEL & FILE LOCKS — you never need to ask the user:
- Excel is automatically killed before every tool call. You do not need to ask the user to close Excel.
- If you see a file lock error, just run compile_check again — it kills Excel first.
- Never ask the user to open Task Manager or close processes.

ask_user — only for genuine decisions:
- Business rules, numbers, or decisions only the user can make.
- NOT for Excel errors, file locks, compile errors, or anything you can fix by reading code.
- NOT when you already know what's wrong from the error text.

SERVER-SIDE FILES — you CANNOT deploy these:
- app.py, ami_optix/*.py, rules_config.yml are SERVER-SIDE files deployed to Render via GitHub.
- This machine does NOT have GitHub access.
- If you find a bug in a server-side file, use flag_server_change to report it.
- The user will coordinate the GitHub push from their dev machine.
- NEVER silently edit server-side files without flagging. The user MUST know.

${AGENT_CONTEXT_TEXT}

CODEBASE:
- VBA: excel-addin/src/*.bas, *.cls
- PowerShell: tools/excel-agent/*.ps1, *.psm1
- Programs: "uap" = UAP; "mih" = MIH (DIFFERENT RULES — see context above)
- Key modules: AMI_Optix_Automation, AMI_Optix_ResultsWriter, AMI_Optix_RentCalcTables, AMI_Optix_Diagnostics, AMI_Optix_VerifyManualRents
- Tests: Diagnostics smoke, Run UAP optimization, Run MIH optimization, Verify Manual Rents (API)`;

// ── Resume / Abort endpoints ──────────────────────────────────────────────────
app.post('/api/resume/:sessionId', (req, res) => {
  const session = sessions.get(req.params.sessionId);
  if (!session || !session.resolvePause) return res.status(404).json({ error: 'No paused session' });
  const answer = (req.body && req.body.answer) ? String(req.body.answer) : '';
  const resolve = session.resolvePause;
  session.resolvePause = null;
  session.rejectPause  = null;
  resolve(answer);
  res.json({ ok: true });
});

app.post('/api/abort/:sessionId', (req, res) => {
  const session = sessions.get(req.params.sessionId);
  if (session) {
    session.aborted = true;
    if (session.activeProcess) { try { session.activeProcess.kill(); } catch {} }
    if (session.rejectPause)   { session.rejectPause(new Error('Aborted by user')); }
  }
  res.json({ ok: true });
});

// ── Chat endpoint (SSE + streaming OpenAI) ────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { message, image, history = [] } = req.body;
  const sessionId = crypto.randomUUID();

  sessions.set(sessionId, { aborted: false, activeProcess: null, resolvePause: null, rejectPause: null });

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.write(`data: ${JSON.stringify({ type: 'session', data: sessionId })}\n\n`);

  const send = (type, data) => {
    if (!res.writableEnded) res.write(`data: ${JSON.stringify({ type, data })}\n\n`);
  };

  res.on('close', () => {
    const s = sessions.get(sessionId);
    if (s) {
      s.aborted = true;
      if (s.activeProcess) { try { s.activeProcess.kill(); } catch {} }
      if (s.rejectPause)   { s.rejectPause(new Error('Client disconnected')); }
    }
  });

  // Merge persisted history with any history sent by client (client wins for recent msgs)
  const persisted = loadHistory();
  const combined = persisted.slice(-40).concat(history.slice(-10));
  // Deduplicate by keeping last occurrence of each role+content pair
  const seen = new Set();
  const dedupedHistory = combined.filter(m => {
    const key = m.role + '|' + (typeof m.content === 'string' ? m.content.slice(0, 80) : '');
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).slice(-20);

  // Build user message — with optional image
  const userContent = image
    ? [{ type: 'text', text: message }, { type: 'image_url', image_url: { url: image } }]
    : message;

  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    ...dedupedHistory,
    { role: 'user', content: userContent }
  ];

  try {
    let iterations = 0;
    let lastError = null;
    let sameErrorStrikes = 0;

    while (iterations < 30) {
      iterations++;
      const session = sessions.get(sessionId);
      if (session && session.aborted) break;

      // ── Streaming OpenAI call ──
      const stream = await openai.chat.completions.create({
        model: 'gpt-5.2',
        messages,
        tools: TOOLS,
        tool_choice: 'auto',
        stream: true
      });

      let assistantContent = '';
      let toolCallsMap = {};
      let finishReason = null;

      for await (const chunk of stream) {
        const session = sessions.get(sessionId);
        if (session && session.aborted) break;

        const delta = chunk.choices[0]?.delta;
        finishReason = chunk.choices[0]?.finish_reason || finishReason;

        // Stream text in real time
        if (delta?.content) {
          assistantContent += delta.content;
          send('text_chunk', delta.content);
        }

        // Accumulate tool call chunks
        if (delta?.tool_calls) {
          for (const tc of delta.tool_calls) {
            if (!toolCallsMap[tc.index]) {
              toolCallsMap[tc.index] = { id: '', type: 'function', function: { name: '', arguments: '' } };
            }
            if (tc.id)                    toolCallsMap[tc.index].id += tc.id;
            if (tc.function?.name)        toolCallsMap[tc.index].function.name += tc.function.name;
            if (tc.function?.arguments)   toolCallsMap[tc.index].function.arguments += tc.function.arguments;
          }
        }
      }

      const toolCalls = Object.values(toolCallsMap);

      // Push assistant message to history
      const assistantMsg = { role: 'assistant', content: assistantContent || null };
      if (toolCalls.length) assistantMsg.tool_calls = toolCalls;
      messages.push(assistantMsg);

      if (finishReason === 'stop' || !toolCalls.length) break;

      // ── Execute tools ──
      for (const toolCall of toolCalls) {
        const session = sessions.get(sessionId);
        if (session && session.aborted) break;

        const toolName = toolCall.function.name;
        let toolArgs = {};
        try { toolArgs = JSON.parse(toolCall.function.arguments || '{}'); } catch {}

        send('tool_start', toolName);
        const result = await executeTool(toolName, toolArgs, sessionId, send);
        send('tool_done', toolName);

        // Same-error escalation (acceptance tests + compile_check)
        if (toolName === 'run_acceptance_tests' || toolName === 'run_full_refresh' || toolName === 'compile_check') {
          const r = String(result);
          const m = r.match(/COMPILE_ERROR[^\n]*/) || r.match(/code-failure[^\n]*/) || r.match(/Variable not defined[^\n]*/);
          const currentError = m ? m[0].slice(0, 200) : null;
          if (currentError && currentError === lastError) {
            sameErrorStrikes++;
            if (sameErrorStrikes >= 2) {
              send('text_chunk', '\n\n⚠️ Same error twice — asking you for guidance.\n');
              const answer = await executeTool('ask_user', {
                question: `I've tried to fix this error twice and it keeps coming back:\n\n"${currentError}"\n\nWhat do you want me to do? You can describe what you see in Excel, ask me to try a different approach, or let me know if you need to do something manually first.`
              }, sessionId, send);
              messages.push({ role: 'tool', tool_call_id: toolCall.id, content: String(result) });
              messages.push({ role: 'user', content: 'User replied: ' + answer });
              sameErrorStrikes = 0;
              lastError = null;
              continue;
            }
          } else {
            lastError = currentError;
            sameErrorStrikes = 0;
          }
        }

        const NO_TRUNCATE = new Set(['read_source_file', 'list_source_files', 'edit_source_file']);
        const resultStr = String(result);
        const content = NO_TRUNCATE.has(toolName) || resultStr.length <= 8000
          ? resultStr
          : resultStr.slice(0, 8000) + '\n[...truncated]';

        messages.push({ role: 'tool', tool_call_id: toolCall.id, content });
      }
    }

    // Save conversation to disk
    try {
      const hist = loadHistory();
      hist.push({ role: 'user', content: typeof userContent === 'string' ? userContent : message, ts: new Date().toISOString() });
      const lastAssistant = [...messages].reverse().find(m => m.role === 'assistant' && typeof m.content === 'string');
      if (lastAssistant) hist.push({ role: 'assistant', content: lastAssistant.content, ts: new Date().toISOString() });
      saveHistory(hist);
    } catch {}
    send('done', null);
  } catch (err) {
    if (!err.message.includes('borted') && !err.message.includes('disconnected')) {
      send('error', err.message);
    } else {
      send('stopped', null);
    }
  } finally {
    sessions.delete(sessionId);
    res.end();
  }
});

app.listen(PORT, () => {
  console.log(`AMI Optix Chat Agent → http://localhost:${PORT}`);
});
