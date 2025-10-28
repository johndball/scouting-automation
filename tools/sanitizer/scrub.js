// tools/sanitizer/scrub.js
#!/usr/bin/env node
const fs=require('fs');const path=require('path');
const rules=[
  {re:/\/d\/[A-Za-z0-9_-]{20,}/g,repl:'/d/REDACTED_ID'},
  {re:/[A-Za-z0-9._%+-]+@group\.calendar\.google\.com/g,repl:'c_example@group.calendar.google.com'},
  {re:/https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec/g,repl:'https://script.google.com/macros/s/DEPLOYMENT_ID/exec'},
  {re:/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g,repl:'someone@example.org'},
  {re:/\b(?:\+?1[-.● ]?)?\(?\d{3}\)?[-.● ]?\d{3}[-.● ]?\d{4}\b/g,repl:'555-0100'},
];
function scrubFile(fp){const o=fs.readFileSync(fp,'utf8');let d=o;for(const{re,repl}of rules)d=d.replace(re,repl);if(d!==o){fs.writeFileSync(fp,d,'utf8');console.log('Sanitized:',fp);}}
function walk(dir){for(const f of fs.readdirSync(dir)){const p=path.join(dir,f);const s=fs.statSync(p);if(s.isDirectory())walk(p);else if(/\.(gs|js|ts|json|md|html|css|yml|yaml|txt)$/i.test(f))scrubFile(p);}}
walk(process.argv[2]||process.cwd());
console.log('Sanitize pass complete.');
