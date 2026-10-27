/* global Office, Word */
const SHAPES = {
  simple: {
    "Star": "   *   \n  ***  \n*******\n  ***  \n   *   ",
    "Arrow": "----->",
    "Heart": " ** ** \n*******\n ***** \n  ***  \n   *   ",
    "Box": "+-----+\n|  +  |\n+-----+"
  },
  mid: {
    "Rocket": "    /\\\n   /  \\\n  |    |\n  | NASA|\n  |    |\n /|____|\\\n   /  \\\n  /____\\",
    "Mountain": "       /\\\n      /  \\\n  /\\ /    \\\n /  V      \\\n/___________\\",
    "Target": "   .----- .\n  /  .-.   \\\n |  ( O )   |\n  \\  '-'  /\n   '-----'",
    "Speech": ".----------------.\n|  say something  |\n'--------.-------'\n         \\__"
  },
  complex: {
    "Castle": "          |>>>                 |>>>\n          |                    |\n      _  _|_  _            _  _|_  _\n     | |_| |_| |          | |_| |_| |\n     |_________| |________|_________|\n     |  _   _   |          |   _   _ |\n     | | | | |  |   FRAME  |  | | | ||\n     |_|_|_|_|__|__________|__|_|_|_||",
    "Dragon": "                 /\\  //\\\n        |\\___/|      /  \\\n        )     (      /    \\\n       =\\     \\    /  /\\  \\\n         )     (   /  /  \\  \\\n        /       \\_/  /    \\  \\\n        \\       /   /      \\  \\\n         \\    _/   /________\\  \\\n          )  /     (          )  )",
    "Robot": "      .---------.\n     /  .---.    \\\n    |  | o o |    |\n    |  |  ^  |    |\n    |  | '-' |    |\n    |   '---'     |\n    |  /|___|\\    |\n    '-------------'",
    "Wave": "        .-~~~-._.-~~~-._.-~~~-._\n   .-~~     ~~~     ~~~     ~~~~-._\n  (      DEEP      OCEAN      )\n   `-._     _.-~~~-._.-~~~-._.-'\n       `~~~'"
  }
};
const $ = id => document.getElementById(id);
const clean = (value, width) => { const text=(value||'').replace(/[\r\n]+/g,' ').trim(); return text.length>width ? text.slice(0,width-1)+'…' : text.padEnd(width,' '); };
function lineChars(style){return {single:['┌','┐','└','┘','─','│'],double:['╔','╗','╚','╝','═','║'],ascii:['+','+','+','+','-','|'],heavy:['┏','┓','┗','┛','━','┃']}[style];}
function drawFrame(){
  const tier=$('drawTier').value, style=$('lineShape').value, chars=lineChars(style), width=tier==='simple'?34:tier==='mid'?44:54, rows=tier==='simple'?5:tier==='mid'?7:9;
  let [tl,tr,bl,br,h,v]=chars;
  const corners={square:[tl,tr,bl,br],round:style==='ascii'?['/','\\','\\','/']:['╭','╮','╰','╯'],diamond:['◇','◇','◇','◇'],star:['✦','✦','✦','✦']}[$('cornerShape').value]; [tl,tr,bl,br]=corners;
  const title=clean($('drawTitle').value||'Drawn title',width-2), body=clean($('drawBody').value||'Choose your shapes',width-2);
  const top=tl+h.repeat(width)+tr, bottom=bl+h.repeat(width)+br, out=[top];
  for(let i=0;i<rows;i++){let inside=i===Math.floor(rows/2)?(i===Math.floor(rows/2)?title:body):i===Math.floor(rows/2)+1?body:' '.repeat(width); out.push(v+inside+v)} out.push(bottom); return out.join('\n');
}
function presetText(){
  const type=$('presetType').value, title=clean($('presetTitle').value||'Frame Studio',type==='asciiFullFrame'?42:52), body=clean($('presetBody').value||'A designed note for Word',type==='asciiFullFrame'?42:52);
  if(type==='headerFrame') return '╔'+ '═'.repeat(52) +'╗\n║ '+title+' ║\n╚'+ '═'.repeat(52) +'╝';
  if(type==='bottomFrame') return '┌'+ '─'.repeat(52) +'┐\n│ '+body+' │\n└'+ '─'.repeat(52) +'┘';
  if(type==='asciiFullFrame'){const w=42, top='+'+'-'.repeat(w)+'+'; return [top,'| '+title+' |','|'+' '.repeat(w+2)+'|','| '+body+' |',...Array(14).fill('|'+' '.repeat(w+2)+'|'),top].join('\n')}
  const w=52, top='╔'+'═'.repeat(w)+'╗'; return [top,'║ '+title+' ║','║ '+body+' ║',...Array(17).fill('║'+' '.repeat(w)+'║'),'╚'+'═'.repeat(w)+'╝'].join('\n');
}
function currentInsertShape(){return SHAPES[$('insertTier').value][$('insertShape').value]||'';}
function refreshPreset(){ $('presetPreview').textContent=presetText(); }
function refreshDraw(){ $('drawPreview').textContent=drawFrame(); }
function populateShapes(){const tier=$('insertTier').value; $('insertShape').innerHTML=Object.keys(SHAPES[tier]).map(name=>`<option>${name}</option>`).join(''); $('insertPreview').textContent=currentInsertShape();}
function notice(message,error=false){const el=$('status');el.textContent=message;el.style.color=error?'#c0395a':'var(--accent)';setTimeout(()=>{if(el.textContent===message)el.textContent=''},3200)}
async function put(text){try{await Word.run(async context=>{const range=context.document.getSelection();range.insertText(text,Word.InsertLocation.end);range.font.name='Consolas';range.font.size=10.5;range.font.color='152033';await context.sync()});notice('Inserted at the cursor.')}catch(e){console.error(e);notice('Could not edit the document. Is Word ready?',true)}}
function copy(text){navigator.clipboard.writeText(text).then(()=>notice('Copied to clipboard.')).catch(()=>notice('Copy failed — select the preview manually.',true));}
function bind(){document.querySelectorAll('.tab').forEach(tab=>tab.addEventListener('click',()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));document.querySelectorAll('.tool-section').forEach(x=>x.classList.remove('active'));tab.classList.add('active');$(tab.dataset.tab).classList.add('active')}));
  $('presetType').addEventListener('change',refreshPreset);$('presetTitle').addEventListener('input',refreshPreset);$('presetBody').addEventListener('input',refreshPreset);$('insertPreset').addEventListener('click',()=>put(presetText()));$('copyPreset').addEventListener('click',()=>copy(presetText()));
  $('insertTier').addEventListener('change',populateShapes);$('insertShape').addEventListener('change',()=>{$('insertPreview').textContent=currentInsertShape()});$('insertShapeBtn').addEventListener('click',()=>put(currentInsertShape()));$('copyInsert').addEventListener('click',()=>copy(currentInsertShape()));
  ['drawTier','cornerShape','lineShape','drawTitle','drawBody'].forEach(id=>$(id).addEventListener('input',refreshDraw));$('insertDraw').addEventListener('click',()=>put(drawFrame()));$('copyDraw').addEventListener('click',()=>copy(drawFrame()));
  populateShapes();refreshPreset();refreshDraw();}
Office.onReady(bind);
