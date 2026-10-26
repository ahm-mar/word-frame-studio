/* global Office, Word */
const PRESETS = {
  frames: {
    "Neon Terminal": { open: "╔══════════════════════════════════╗", mid: "║ {title:<32} ║", body: "║ {body:<32} ║", close: "╚══════════════════════════════════╝" },
    "Soft Card": { open: "┌──────────────────────────────────┐", mid: "│ {title:<32} │", body: "│ {body:<32} │", close: "└──────────────────────────────────┘" },
    "Double Line": { open: "╓──────────────────────────────────╖", mid: "║ {title:<32} ║", body: "║ {body:<32} ║", close: "╙──────────────────────────────────╜" },
    "Corner Brackets": { open: "⌜────────────────────────────────⌝", mid: "│ {title:<32} │", body: "│ {body:<32} │", close: "⌞────────────────────────────────⌟" }
  },
  shapes: {
    "Spark Burst": { fixed: "          .     *     .\n      *      .-.-.      *\n   .       /  ✦  \\       .\n      *   |  WOW! |   *\n   .       \\  ✦  /       .\n      *      '-.-'      *\n          .     *     ." },
    "Arrow": { fixed: "                 ╱│\n                ╱ │\n               ╱  │\n──────────────▶   │\n               ╲  │\n                ╲ │\n                 ╲│" },
    "Diamond": { fixed: "             /\\\n            /  \\\n           /    \\\n          <  ✦   >\n           \\    /\n            \\  /\n             \\/" },
    "Coffee": { fixed: "          )  (\n         (    )\n          )  (\n       .-------.\n       |       |\n       |  ☕   |\n       '-------'" }
  },
  dividers: {
    "Signal": { fixed: "─────────────── ✦ ───────────────" },
    "Dots": { fixed: "· · · · · · · · · · · · · · · · · ·" },
    "Wave": { fixed: "~~~~~~~  ~~~~~~~  ~~~~~~~  ~~~~~~~" },
    "Minimal": { fixed: "──────────────────────────────────" }
  }
};

let category = 'frames';
const $ = id => document.getElementById(id);
const fit = (value, width) => {
  const clean = (value || '').replace(/[\r\n]+/g, ' ').trim();
  return clean.length > width ? clean.slice(0, width - 1) + '…' : clean.padEnd(width, ' ');
};
function renderPreset() {
  const preset = PRESETS[category][$('preset').value];
  if (!preset) return '';
  if (preset.fixed) return preset.fixed;
  const title = fit($('titleInput').value || 'Your title', 32);
  const body = fit($('bodyInput').value || 'Your message', 32);
  return [preset.open, preset.mid.replace('{title:<32}', title), preset.body.replace('{body:<32}', body), preset.close].join('\n');
}
function refresh() { $('preview').textContent = renderPreset(); }
function populate() {
  $('preset').innerHTML = Object.keys(PRESETS[category]).map(name => `<option>${name}</option>`).join('');
  const isTextFrame = category === 'frames';
  $('titleInput').disabled = !isTextFrame; $('bodyInput').disabled = !isTextFrame;
  $('titleInput').style.opacity = isTextFrame ? 1 : .5; $('bodyInput').style.opacity = isTextFrame ? 1 : .5;
  refresh();
}
function notify(message, error=false) { const el=$('status'); el.textContent=message; el.style.color=error?'#c0395a':'var(--accent)'; setTimeout(()=>{if(el.textContent===message)el.textContent=''},3200); }
async function insert(replace) {
  const text = renderPreset();
  try {
    await Word.run(async context => {
      const range = context.document.getSelection();
      range.insertText(text, replace ? Word.InsertLocation.replace : Word.InsertLocation.end);
      range.font.name = 'Consolas'; range.font.size = 10.5; range.font.color = '152033';
      await context.sync();
    });
    notify(replace ? 'Selection replaced.' : 'Inserted at the cursor.');
  } catch (e) { notify('Could not edit the document. Is Word ready?', true); console.error(e); }
}
$('category').addEventListener('change', e => { category=e.target.value; populate(); });
$('preset').addEventListener('change', refresh); $('titleInput').addEventListener('input', refresh); $('bodyInput').addEventListener('input', refresh);
$('insertBtn').addEventListener('click', () => insert(false)); $('replaceBtn').addEventListener('click', () => insert(true));
$('copyBtn').addEventListener('click', async () => { try { await navigator.clipboard.writeText(renderPreset()); notify('Copied to clipboard.'); } catch { notify('Copy failed — select the preview manually.', true); } });
Office.onReady(() => populate());
