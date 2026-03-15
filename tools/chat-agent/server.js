'use strict';

const express = require('express');
const OpenAI = require('openai');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const AGENT_ROOT = process.env.AGENT_ROOT || 'C:\\AMI_Optix_Agent';
const REPO_ROOT  = process.env.REPO_ROOT  || '';
const PORT       = parseInt(process.env.PORT || '3000', 10);

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY || '' });

// ── Tool definitions ────────────────────────────────────────────────────────

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
            description: 'Path relative to repo root, e.g. excel-addin/src/AMI_Optix_RentCalcTables.bas'
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
      description: 'Make a targeted edit to a source file by replacing an exact string. The old_string must match exactly (including whitespace).',
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
  }
];

// ── Tool implementations ─────────────────────────────────────────────────────

function runPowershell(scriptPath, args) {
  return new Promise((resolve) => {
    const ps = spawn('powershell', [
      '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, ...args
    ]);
    let stdout = '';
    let stderr = '';
    ps.stdout.on('data', d => { stdout += d.toString(); });
    ps.stderr.on('data', d => { stderr += d.toString(); });
    ps.on('close', code => resolve({ code, stdout, stderr }));
  });
}

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

async function executeTool(name, args) {
  switch (name) {

    case 'run_acceptance_tests': {
      const script = path.join(AGENT_ROOT, 'scripts', 'Invoke-AmiOptixAcceptance.ps1');
      if (!fs.existsSync(script)) {
        return 'Acceptance script not found. Run a full refresh first to bootstrap the workspace.';
      }
      const r = await runPowershell(script, ['-AgentRoot', AGENT_ROOT]);
      return (r.stdout + (r.stderr ? '\nSTDERR: ' + r.stderr : '')).trim();
    }

    case 'run_full_refresh': {
      const repoScript = REPO_ROOT
        ? path.join(REPO_ROOT, 'tools', 'excel-agent', 'Refresh-AmiOptixAgent.ps1')
        : path.join(AGENT_ROOT, 'scripts', 'Refresh-AmiOptixAgent.ps1');
      const psArgs = ['-AgentRoot', AGENT_ROOT];
      if (REPO_ROOT) psArgs.push('-RepoRoot', REPO_ROOT);
      const r = await runPowershell(repoScript, psArgs);
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
      try {
        return fs.readFileSync(filePath, 'utf8');
      } catch (e) {
        return `Error reading file: ${e.message}`;
      }
    }

    case 'edit_source_file': {
      if (!REPO_ROOT) return 'REPO_ROOT is not configured.';
      const filePath = path.join(REPO_ROOT, args.relative_path.replace(/\//g, path.sep));
      try {
        const content = fs.readFileSync(filePath, 'utf8');
        if (!content.includes(args.old_string)) {
          return 'ERROR: old_string not found in file — no changes made. Double-check whitespace and line endings.';
        }
        const updated = content.replace(args.old_string, args.new_string);
        fs.writeFileSync(filePath, updated, 'utf8');
        return `OK: edited ${args.relative_path}`;
      } catch (e) {
        return `Error editing file: ${e.message}`;
      }
    }

    case 'read_results': {
      const p = path.join(AGENT_ROOT, 'artifacts', 'acceptance-result.json');
      try { return fs.readFileSync(p, 'utf8'); }
      catch (e) { return `No acceptance results found: ${e.message}`; }
    }

    case 'read_build_result': {
      const p = path.join(AGENT_ROOT, 'artifacts', 'build-result.json');
      try { return fs.readFileSync(p, 'utf8'); }
      catch (e) { return `No build results found: ${e.message}`; }
    }

    default:
      return `Unknown tool: ${name}`;
  }
}

// ── System prompt ─────────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `You are an autonomous repair agent for AMI Optix — a NYC affordable housing AMI (Area Median Income) calculator implemented as a VBA Excel add-in with a PowerShell automation layer.

When the user describes a problem, you work autonomously through this loop:
1. Run acceptance tests to see what is currently failing
2. Read the relevant source files to find the root cause
3. Edit the source files to fix the issue
4. Run a full refresh (rebuild + retest)
5. Loop until all tests pass (max 5 fix attempts)
6. Report back with a plain-English summary: what was broken, what you changed, and the final test results

Rules:
- Do NOT ask the user clarifying questions unless you genuinely cannot proceed without information only they know
- Do NOT ask for approval before editing files — just do it
- ONLY pause and ask the user if the fix requires a security-sensitive change (credentials, auth, data access permissions)
- When all tests pass, stop working and report success clearly
- If you try 5 fix attempts and tests still fail, summarize what you found and what you tried, and tell the user what you need from them

The codebase:
- VBA source files: excel-addin/src/*.bas, *.cls
- PowerShell agent: tools/excel-agent/*.ps1, *.psm1
- Acceptance config: tools/excel-agent/config/acceptance.template.json
- Two workbook roles: "uap" (utility allowance program) and "mih" (mandatory inclusionary housing)
- Key modules: AMI_Optix_Automation (optimization), AMI_Optix_VerifyManualRents (rent verification), AMI_Optix_RentCalcTables (rent calculations), AMI_Optix_ResultsWriter (results output), AMI_Optix_Diagnostics (diagnostics sheet)
- Acceptance tests: Diagnostics smoke, Run UAP optimization, Run MIH optimization, Verify Manual Rents (API)`;

// ── Chat endpoint (SSE) ───────────────────────────────────────────────────────

app.post('/api/chat', async (req, res) => {
  const { message, history = [] } = req.body;

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const send = (type, data) => {
    res.write(`data: ${JSON.stringify({ type, data })}\n\n`);
  };

  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    ...history,
    { role: 'user', content: message }
  ];

  try {
    let iterations = 0;
    const MAX_ITERATIONS = 30;

    while (iterations < MAX_ITERATIONS) {
      iterations++;

      const response = await openai.chat.completions.create({
        model: 'gpt-4o',
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
        const toolName = toolCall.function.name;
        let toolArgs = {};
        try { toolArgs = JSON.parse(toolCall.function.arguments || '{}'); } catch {}

        send('tool_start', toolName);

        const result = await executeTool(toolName, toolArgs);

        send('tool_done', toolName);

        messages.push({
          role: 'tool',
          tool_call_id: toolCall.id,
          content: String(result)
        });
      }
    }

    send('done', null);
  } catch (err) {
    send('error', err.message);
  }

  res.end();
});

app.listen(PORT, () => {
  console.log(`AMI Optix Chat Agent running → http://localhost:${PORT}`);
});
