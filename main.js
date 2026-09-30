const { Plugin, Notice } = require("obsidian");
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");

module.exports = class OneNoteVaultSyncPlugin extends Plugin {
  async onload() {
    this.addRibbonIcon("refresh-cw", "OneNote: Sync Vault", () => {
      this.syncVault();
    });

    this.addCommand({
      id: "sync-vault",
      name: "OneNote: Sync Vault",
      callback: () => this.syncVault(),
    });
  }

  getScriptPath(vaultPath) {
    const candidates = [
      path.join(vaultPath, this.manifest.dir || "", "scripts", "sync-onenote.cmd"),
      path.join(vaultPath, ".obsidian", "plugins", this.manifest.id, "scripts", "sync-onenote.cmd"),
      path.join(vaultPath, "copilot", "scripts", "sync-onenote.cmd"),
      path.join(vaultPath, "scripts", "sync-onenote.cmd"),
    ];

    for (const candidate of candidates) {
      if (candidate && fs.existsSync(candidate)) {
        return candidate;
      }
    }
    return candidates[0];
  }

  syncVault() {
    const vaultPath = this.app.vault.adapter.getBasePath();
    const scriptPath = this.getScriptPath(vaultPath);

    if (!fs.existsSync(scriptPath)) {
      new Notice(`OneNote sync script not found: ${scriptPath}`, 10000);
      return;
    }

    new Notice("OneNote: Checking for changes...", 3000);
    const process = spawn("cmd.exe", ["/c", scriptPath], {
      cwd: vaultPath,
      windowsHide: true,
    });

    let output = "";
    let errorOutput = "";

    process.stdout.on("data", (data) => {
      output += data.toString();
    });

    process.stderr.on("data", (data) => {
      errorOutput += data.toString();
    });

    process.on("error", (error) => {
      new Notice(`OneNote sync could not be started: ${error.message}`, 10000);
    });

    process.on("close", (code) => {
      if (code === 0) {
        const lines = output.trim().split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
        const lastLine = lines.length > 0 ? lines[lines.length - 1] : "OneNote is up to date.";
        new Notice(lastLine, 6000);
      } else {
        const detail = (errorOutput.trim() || output.trim());
        new Notice(`OneNote sync failed (${code ?? "unknown"}).${detail ? ` ${detail}` : ""}`, 10000);
      }
    });
  }
};
