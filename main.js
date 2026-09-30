const { Plugin, Notice, PluginSettingTab, Setting } = require("obsidian");
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");

const DEFAULT_SETTINGS = {
  createCanvas: true,
  skipInkImages: false,
  notebookFilter: "",
  sectionFilter: "",
};

class OneNoteSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();

    containerEl.createEl("h2", { text: "OneNote Vault Sync Settings" });

    new Setting(containerEl)
      .setName("Create Canvas View")
      .setDesc("Automatically generate an Obsidian .canvas file for notes with spatial or drawing layouts.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.createCanvas)
          .onChange(async (value) => {
            this.plugin.settings.createCanvas = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Skip Ink Drawing Images")
      .setDesc("Do not render handwriting/drawing pen strokes as PNG images (recognized text will still be exported).")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.skipInkImages)
          .onChange(async (value) => {
            this.plugin.settings.skipInkImages = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Notebook Filter")
      .setDesc("Optional: Only synchronize a specific notebook (leave empty for all).")
      .addText((text) =>
        text
          .setPlaceholder("e.g. My Notebook")
          .setValue(this.plugin.settings.notebookFilter)
          .onChange(async (value) => {
            this.plugin.settings.notebookFilter = value.trim();
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Section Filter")
      .setDesc("Optional: Only synchronize a specific section (leave empty for all).")
      .addText((text) =>
        text
          .setPlaceholder("e.g. Quick Notes")
          .setValue(this.plugin.settings.sectionFilter)
          .onChange(async (value) => {
            this.plugin.settings.sectionFilter = value.trim();
            await this.plugin.saveSettings();
          })
      );
  }
}

module.exports = class OneNoteVaultSyncPlugin extends Plugin {
  async onload() {
    await this.loadSettings();

    this.addRibbonIcon("refresh-cw", "OneNote: Sync Vault", () => {
      this.syncVault();
    });

    this.addCommand({
      id: "sync-vault",
      name: "OneNote: Sync Vault",
      callback: () => this.syncVault(),
    });

    this.addSettingTab(new OneNoteSettingTab(this.app, this));
  }

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings() {
    await this.saveData(this.settings);
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

    const args = ["/c", scriptPath];
    if (!this.settings.createCanvas) {
      args.push("-SkipCanvas");
    }
    if (this.settings.skipInkImages) {
      args.push("-SkipInkImages");
    }
    if (this.settings.notebookFilter) {
      args.push("-Notebook", `"${this.settings.notebookFilter}"`);
    }
    if (this.settings.sectionFilter) {
      args.push("-Section", `"${this.settings.sectionFilter}"`);
    }

    new Notice("OneNote: Checking for changes...", 3000);
    const process = spawn("cmd.exe", args, {
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
