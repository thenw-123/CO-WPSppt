import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import { promisify } from 'util';

const execFileAsync = promisify(cp.execFile);

let output: vscode.OutputChannel;

function getBundledRoot(context: vscode.ExtensionContext): string {
  return path.join(context.extensionPath, 'bundled');
}

function getDispatchScript(bundled: string): string {
  return path.join(bundled, 'tools', 'wps_dispatch.ps1');
}

async function runDispatch(
  context: vscode.ExtensionContext,
  action: string,
  argsJson: string
): Promise<{ stdout: string; stderr: string }> {
  const cfg = vscode.workspace.getConfiguration('wpsPpt');
  const ps = cfg.get<string>('powershellExecutable') || 'powershell.exe';
  const pol = cfg.get<string>('executionPolicy') || 'Bypass';
  const pretty = cfg.get<boolean>('prettyJson') ? '1' : '0';

  const bundled = getBundledRoot(context);
  const script = getDispatchScript(bundled);
  if (!(await fileExists(script))) {
    throw new Error(
      `Bundled script missing: ${script}. Run: npm run bundle (in extension/) or clone full repo and build VSIX with vscode:prepublish.`
    );
  }

  const env = { ...process.env, OC_PRETTY_JSON: pretty };
  const { stdout, stderr } = await execFileAsync(
    ps,
    ['-NoProfile', '-ExecutionPolicy', pol, '-File', script, '-Action', action, '-ArgsJson', argsJson],
    {
      env,
      windowsHide: true,
      maxBuffer: 10 * 1024 * 1024,
    }
  );
  return { stdout: stdout.trim(), stderr: stderr.trim() };
}

async function fileExists(p: string): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(vscode.Uri.file(p));
    return true;
  } catch {
    return false;
  }
}

function logResult(stdout: string, stderr: string) {
  output.appendLine('--- stdout ---');
  output.appendLine(stdout || '(empty)');
  if (stderr) {
    output.appendLine('--- stderr ---');
    output.appendLine(stderr);
  }
  output.show(true);
}

export function activate(context: vscode.ExtensionContext) {
  output = vscode.window.createOutputChannel('WPS PPT');

  const reg = (cmd: string, fn: () => void | Promise<void>) => {
    context.subscriptions.push(vscode.commands.registerCommand(cmd, fn));
  };

  reg('wpsPpt.doctor', async () => {
    try {
      const { stdout, stderr } = await runDispatch(context, 'doctor', '{}');
      logResult(stdout, stderr);
      const ok = stdout.includes('"ok":true');
      vscode.window.showInformationMessage(ok ? 'WPS PPT: doctor finished (see output).' : 'WPS PPT: doctor reported issues (see output).');
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      output.appendLine(msg);
      output.show(true);
      vscode.window.showErrorMessage(`WPS PPT doctor failed: ${msg}`);
    }
  });

  reg('wpsPpt.validateSpec', async () => {
    const specPath = await pickSpecFile();
    if (!specPath) {
      return;
    }
    const args = JSON.stringify({ specPath: toBundledOrAbsoluteSpecPath(specPath, context) });
    try {
      const { stdout, stderr } = await runDispatch(context, 'validate-spec', args);
      logResult(stdout, stderr);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      vscode.window.showErrorMessage(msg);
    }
  });

  reg('wpsPpt.runSpec', async () => {
    const specPath = await pickSpecFile();
    if (!specPath) {
      return;
    }
    const args = JSON.stringify({ specPath: toBundledOrAbsoluteSpecPath(specPath, context) });
    try {
      const { stdout, stderr } = await runDispatch(context, 'run-spec', args);
      logResult(stdout, stderr);
      const ok = stdout.includes('"ok":true');
      if (ok) {
        vscode.window.showInformationMessage('WPS PPT: run-spec finished. Check output for presentationPath.');
      } else {
        vscode.window.showWarningMessage('WPS PPT: run-spec reported failure. See output.');
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      vscode.window.showErrorMessage(msg);
    }
  });

  reg('wpsPpt.runSpecExample', async () => {
    const args = JSON.stringify({ specPath: 'specs/example-spec.json' });
    try {
      const { stdout, stderr } = await runDispatch(context, 'run-spec', args);
      logResult(stdout, stderr);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      vscode.window.showErrorMessage(msg);
    }
  });

  reg('wpsPpt.openManifest', async () => {
    const p = path.join(getBundledRoot(context), 'manifest.json');
    await openFile(p);
  });

  reg('wpsPpt.openSkill', async () => {
    const p = path.join(getBundledRoot(context), 'skills', 'drive-wps-ppt', 'SKILL.md');
    await openFile(p);
  });

  reg('wpsPpt.openBundledSpecs', async () => {
    const p = path.join(getBundledRoot(context), 'specs');
    await vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(p));
  });
}

/**
 * If user picks a file inside bundled/specs, pass relative path from bundled root;
 * otherwise pass absolute path (run-spec supports rooted specPath).
 */
function toBundledOrAbsoluteSpecPath(fsPath: string, context: vscode.ExtensionContext): string {
  const bundled = getBundledRoot(context);
  const normalized = path.normalize(fsPath);
  if (normalized.toLowerCase().startsWith(bundled.toLowerCase() + path.sep)) {
    return path.relative(bundled, normalized).split(path.sep).join('/');
  }
  return normalized;
}

async function pickSpecFile(): Promise<string | undefined> {
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: 'Select spec JSON',
    filters: { JSON: ['json'] },
  });
  if (!picked?.[0]) {
    return undefined;
  }
  return picked[0].fsPath;
}

async function openFile(fsPath: string): Promise<void> {
  if (!(await fileExists(fsPath))) {
    vscode.window.showErrorMessage(`File not found: ${fsPath}. Build the extension (bundle) first.`);
    return;
  }
  const doc = await vscode.workspace.openTextDocument(fsPath);
  await vscode.window.showTextDocument(doc);
}

export function deactivate() {
  output?.dispose();
}
