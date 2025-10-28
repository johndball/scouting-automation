# tools/sanitizer/scrub.py
import re, sys, os
RULES=[
 (re.compile(r"/d/[A-Za-z0-9_-]{20,}"),"/d/REDACTED_ID"),
 (re.compile(r"[A-Za-z0-9._%+-]+@group\.calendar\.google\.com"),"c_example@group.calendar.google.com"),
 (re.compile(r"https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec"),"https://script.google.com/macros/s/DEPLOYMENT_ID/exec"),
 (re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"),"someone@example.org"),
 (re.compile(r"\b(?:\+?1[-.● ]?)?\(?\d{3}\)?[-.● ]?\d{3}[-.● ]?\d{4}\b"),"555-0100"),
]
EXTS=(".gs",".js",".ts",".json",".md",".html",".css",".yml",".yaml",".txt")

def scrub_file(p):
    with open(p,'r',encoding='utf-8',errors='ignore') as f: data=f.read()
    orig=data
    for rx,repl in RULES: data=rx.sub(repl,data)
    if data!=orig:
        with open(p,'w',encoding='utf-8') as f: f.write(data)
        print('Sanitized:',p)

def walk(root):
    for dp,_,fns in os.walk(root):
        for fn in fns:
            if fn.lower().endswith(EXTS): scrub_file(os.path.join(dp,fn))

walk(sys.argv[1] if len(sys.argv)>1 else os.getcwd())
print('Sanitize pass complete.')
