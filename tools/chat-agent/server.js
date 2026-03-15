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
// sessionId → { aborted, activeProcess, resolvePause, rejectPause }
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
          relative_path: {
            type: 'string',
            description: 'Path relative to repo root, e.g. excel-addin/src/AMI_Optix_ResultsWriter.bas'
          }
        },
        required: ['relative_path']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'edit_source_file',
      description: 'Make a targeted edit to a source file by replacing an exact string. The old_string must match exactly (including whitespace and line endings).',
      parameters: {
        type: 'object',
        properties: {
          relative_path: { type: 'string', description: 'Path relative to repo root' },
          old_string:    { type: 'string', description: 'Exact text to replace' },
          new_string:    { type: 'string', description: 'Replacement text' }
        },
        required: ['relative_path', 'old_string', 'new_string']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_results',
      description: 'Read the latest acceptance test results JSON from the agent workspace.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_build_result',
      description: 'Read the latest build result JSON from the agent workspace.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'ask_user',
      description: 'Pause and ask the user a question when you are genuinely stuck and cannot proceed without their input. Use this sparingly — only when you truly cannot determine the next step on your own.',
      parameters: {
        type: 'object',
        properties: {
          question: { type: 'string', description: 'The specific question to ask the user' }
        },
        required: ['question']
      }
    }
  }
];

// ── PowerShell runner with timeout + kill-on-abort ───────────────────────────
function runPowershell(scriptPath, args, sessionId, timeoutMs = 300000) {
  return new Promise((resolve) => {
    const ps = spawn('powershell', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, ...args
    ]);

    const session = sessions.get(sessionId);
    if (session) session.activeProcess = ps;

    let stdout = '';
    let stderr = '';
    let finished = false;

    const finish = (code) => {
      if (finished) return;
      finished = true;
      if (session) session.activeProcess = null;
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
          stderr: (stderr || '') + '\n[TIMED OUT after 5 minutes — Excel may have a compile error dialog open. Please close Excel manually and try again.]'
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
      if (!fs.existsSync(script)) return 'Acceptance script not found at ' + script + '. Run a full refresh first.';
      const r = await runPowershell(script, ['-AgentRoot', AGENT_ROOT], sessionId);
      return (r.stdout + (r.stderr ? '\nSTDERR: ' + r.stderr : '')).trim();
    }

    case 'run_full_refresh': {
      const repoScript = REPO_ROOT
        ? path.join(REPO_ROOT, 'tools', 'excel-agent', 'Refresh-AmiOptixAgent.ps1')
        : path.join(AGENT_ROOT, 'scripts', 'Refresh-AmiOptixAgent.ps1');
      const psArgs = ['-AgentRoot', AGENT_ROOT];
      if (REPO_ROOT) psArgs.push('-RepoRoot', REPO_ROOT);
      const r = await runPowershell(repoScript, psArgs, sessionId);
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
          return 'ERROR: old_string not found in file — no changes made. Check whitespace and line endings exactly.';
        }
        const updated = content.replace(args.old_string, args.new_string);
        fs.writeFileSync(filePath, updated, 'utf8');
        // Send a richer progress event showing what changed
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

When the user describes a problem, work autonomously:
1. Run acceptance tests to see what is failing
2. Read relevant source files to diagnose the root cause
3. Edit source files to fix the issue
4. Run a full refresh (rebuild + retest)
5. If the same error appears after a fix attempt, call ask_user — do NOT keep retrying the same approach
6. When all tests pass, report success clearly

Rules:
- Do NOT ask clarifying questions unless you cannot proceed without information only the user knows
- Call ask_user when you hit a blocker you cannot resolve (same error repeating, ambiguous requirements, Excel hung)
- When all tests pass, stop and report what you fixed
- Never make the same edit twice if it didn't work — try a different approach or ask_user

The codebase:
- VBA source: excel-addin/src/*.bas, *.cls
- PowerShell agent: tools/excel-agent/*.ps1, *.psm1
- Acceptance config: tools/excel-agent/config/acceptance.template.json
- Workbook roles: "uap" (utility allowance program), "mih" (mandatory inclusionary housing)
- Key modules: AMI_Optix_Automation, AMI_Optix_VerifyManualRents, AMI_Optix_RentCalcTables, AMI_Optix_ResultsWriter, AMI_Optix_Diagnostics
- Acceptance tests: Diagnostics smoke, Run UAP optimization, Run MIH optimization, Verify Manual Rents (API)`;

// ── Resume endpoint ───────────────────────────────────────────────────────────
app.post('/api/resume/:sessionId', (req, res) => {
  const session = sessions.get(req.params.sessionId);
  if (!session || !session.resolvePause) {
    return res.status(404).json({ error: 'No paused session found' });
  }
  const answer = (req.body && req.body.answer) ? String(req.body.answer) : '';
  const resolve = session.resolvePause;
  session.resolvePause = null;
  session.rejectPause  = null;
  resolve(answer);
  res.json({ ok: true });
});

// ── Abort endpoint ────────────────────────────────────────────────────────────
app.post('/api/abort/:sessionId', (req, res) => {
  const session = sessions.get(req.params.sessionId);
  if (session) {
    session.aborted = true;
    if (session.activeProcess) {
      try { session.activeProcess.kill(); } catch {}
    }
    if (session.rejectPause) {
      session.rejectPause(new Error('Aborted by user'));
    }
  }
  res.json({ ok: true });
});

// ── Chat endpoint (SSE) ───────────────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { message, history = [] } = req.body;
  const sessionId = crypto.randomUUID();

  sessions.set(sessionId, {
    aborted: false,
    activeProcess: null,
    resolvePause: null,
    rejectPause: null
  });

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  // Send session ID first so client can use it for abort/resume
  res.write(`data: ${JSON.stringify({ type: 'session', data: sessionId })}\n\n`);

  const send = (type, data) => {
    if (!res.writableEnded) {
      res.write(`data: ${JSON.stringify({ type, data })}\n\n`);
    }
  };

  // Kill process and mark aborted when client disconnects
  res.on('close', () => {
    const session = sessions.get(sessionId);
    if (session) {
      session.aborted = true;
      if (session.activeProcess) {
        try { session.activeProcess.kill(); } catch {}
      }
      if (session.rejectPause) {
        session.rejectPause(new Error('Client disconnected'));
      }
    }
  });

  const recentHistory = history.slice(-10);
  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    ...recentHistory,
    { role: 'user', content: message }
  ];

  try {
    let iterations = 0;
    const MAX_ITERATIONS = 30;
    let lastError = null;
    let sameErrorStrikes = 0;

    while (iterations < MAX_ITERATIONS) {
      iterations++;

      const session = sessions.get(sessionId);
      if (session && session.aborted) break;

      const response = await openai.chat.completions.create({
        model: 'gpt-5.2',
        messages,
        tools: TOOLS,
        tool_choice: 'auto'
      });

      const choice = response.choices[0];
      const assistantMessage = choice.message;
      messages.push(assistantMessage);

      if (assistantMessage.content) {
        send('text', assistantMessage.content);
      }

      if (choice.finish_reason === 'stop' || !assistantMessage.tool_calls?.length) {
        break;
      }

      for (const toolCall of assistantMessage.tool_calls) {
        const session = sessions.get(sessionId);
        if (session && session.aborted) break;

        const toolName = toolCall.function.name;
        let toolArgs = {};
        try { toolArgs = JSON.parse(toolCall.function.arguments || '{}'); } catch {}

        send('tool_start', toolName);

        const result = await executeTool(toolName, toolArgs, sessionId, send);

        send('tool_done', toolName);

        // Detect same error repeating — escalate to user after 2 strikes
        if (toolName === 'run_acceptance_tests' || toolName === 'run_full_refresh') {
          const errorMatch = String(result).match(/code-failure[^\n]*/);
          const currentError = errorMatch ? errorMatch[0] : null;
          if (currentError && currentError === lastError) {
            sameErrorStrikes++;
            if (sameErrorStrikes >= 2) {
              send('text', '⚠️ I\'ve hit the same error twice and my last fix didn\'t work. Asking you for guidance...');
              const answer = await executeTool('ask_user', {
                question: `I tried to fix this error twice but it keeps coming back:\n\n"${currentError}"\n\nWhat would you like me to do? You can describe what you see in Excel, tell me to try a different approach, or let me know if you need to do something manually first.`
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

        // Only truncate log output, never source file contents
        const NO_TRUNCATE = new Set(['read_source_file', 'list_source_files', 'edit_source_file']);
        const MAX_LOG = 8000;
        const resultStr = String(result);
        const content = NO_TRUNCATE.has(toolName) || resultStr.length <= MAX_LOG
          ? resultStr
          : resultStr.slice(0, MAX_LOG) + '\n[...truncated]';

        messages.push({
          role: 'tool',
          tool_call_id: toolCall.id,
          content
        });
      }
    }

    send('done', null);
  } catch (err) {
    if (!err.message.includes('aborted') && !err.message.includes('Aborted')) {
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
  console.log(`AMI Optix Chat Agent running → http://localhost:${PORT}`);
});
