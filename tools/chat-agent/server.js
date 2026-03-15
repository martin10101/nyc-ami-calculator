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
      const r = await runPowershell(script, ['-AgentRoot', AGENT_ROOT], sessionId, 60000, send);
      return (r.stdout + (r.stderr ? '\nSTDERR: ' + r.stderr : '')).trim();
    }

    case 'compile_check': {
      const script = path.join(AGENT_ROOT, 'scripts', 'Invoke-AmiOptixCompileCheck.ps1');
      if (!fs.existsSync(script)) return 'Compile check script not found. Run a full refresh first to bootstrap the workspace.';
      const r = await runPowershell(script, ['-AgentRoot', AGENT_ROOT], sessionId, 60000, send);
      return (r.stdout + (r.stderr ? '\nSTDERR: ' + r.stderr : '')).trim();
    }

    case 'run_full_refresh': {
      const repoScript = REPO_ROOT
        ? path.join(REPO_ROOT, 'tools', 'excel-agent', 'Refresh-AmiOptixAgent.ps1')
        : path.join(AGENT_ROOT, 'scripts', 'Refresh-AmiOptixAgent.ps1');
      const psArgs = ['-AgentRoot', AGENT_ROOT];
      if (REPO_ROOT) psArgs.push('-RepoRoot', REPO_ROOT);
      const r = await runPowershell(repoScript, psArgs, sessionId, 60000, send);
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

    default:
      return 'Unknown tool: ' + name;
  }
}

// ── System prompt ─────────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are an autonomous repair agent for AMI Optix — a NYC affordable housing AMI calculator implemented as a VBA Excel add-in with a PowerShell automation layer.

IMPORTANT — narrate every step out loud as you work. Before each action write one sentence: what you are about to do and why. After each tool result write one sentence: what you found. The user must never be in the dark about what you are doing.

Example narration style:
"I'm going to run the acceptance tests to see what is currently failing."
"Tests show the MIH optimization is failing with a compile error in WriteMihSquareFootageSummary — that function is missing."
"I'm reading AMI_Optix_ResultsWriter.bas to find where to add the missing function."
"Found the call site at line 420. I'm going to add the function definition above it."
"Edit applied. Running a full refresh to verify the fix."

When the test result includes '[Excel dialog auto-dismissed. Error text: ...]', that text tells you exactly what compile error Excel threw — use it to diagnose the problem.

Autonomous work loop:
1. Narrate → run acceptance tests to see what is failing
2. Narrate → read relevant source files to find root cause
3. Narrate → edit source files to fix
4. Narrate → call compile_check (fast: rebuilds + verifies compile, ~30s) — fix any COMPILE_ERROR before proceeding
5. Narrate → run full refresh only once compile is clean
6. If same error repeats after a fix attempt → call ask_user (do NOT keep retrying)
7. When all tests pass → report success with a clear summary

IMPORTANT: Always call compile_check after editing VBA source files. Do not call run_full_refresh until compile_check returns COMPILE_OK. This catches missing functions, undefined variables, and syntax errors in seconds instead of minutes.

Rules:
- Never make the same edit twice if it failed — try a different approach or ask_user
- Call ask_user only when you truly cannot proceed without information only the user can provide
- When tests pass, stop and summarize what was broken and what you fixed

Codebase:
- VBA source: excel-addin/src/*.bas, *.cls
- PowerShell agent: tools/excel-agent/*.ps1, *.psm1
- Workbook roles: "uap", "mih"
- Key modules: AMI_Optix_Automation, AMI_Optix_VerifyManualRents, AMI_Optix_RentCalcTables, AMI_Optix_ResultsWriter, AMI_Optix_Diagnostics
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
  const { message, history = [] } = req.body;
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

  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    ...history.slice(-10),
    { role: 'user', content: message }
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

        // Same-error escalation
        if (toolName === 'run_acceptance_tests' || toolName === 'run_full_refresh') {
          const m = String(result).match(/code-failure[^\n]*/);
          const currentError = m ? m[0] : null;
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
