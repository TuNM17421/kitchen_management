# Harness: chay macro VBA trong LibreOffice (che do tuong thich VBA)
import uno, subprocess, time, os, sys, shutil
from com.sun.star.beans import PropertyValue

def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

PORT = 2002
def start():
    subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '--nologo',
                      f'--accept=socket,host=localhost,port={PORT};urp;'],
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(60):
        try:
            ctx = res.resolve(f'uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext')
            return ctx
        except Exception:
            time.sleep(1)
    raise RuntimeError('no soffice')

def desktop(ctx):
    return ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

def load(ctx, path):
    url = uno.systemPathToFileUrl(os.path.abspath(path))
    return desktop(ctx).loadComponentFromURL(url, '_blank', 0, (pv('Hidden', False), pv('MacroExecutionMode', 4)))

def vba_src(path, is_class=False):
    s = open(path, encoding='latin-1', newline='').read().replace('\r\n', '\n')
    lines = s.split('\n')
    out = []
    skip_hdr = is_class
    for ln in lines:
        if skip_hdr:
            if ln.startswith('VERSION ') or ln.startswith('BEGIN') or ln.startswith('  MultiUse') or ln.startswith('END'):
                continue
            skip_hdr = False if not ln.startswith('Attribute') and ln.strip() != '' and not ln.startswith('END') else skip_hdr
        if ln.startswith('Attribute VB_'):
            continue
        out.append(ln)
    head = 'Option VBASupport 1\n' + ('Option ClassModule\n' if is_class else '')
    return head + '\n'.join(out)

def set_modules(doc, mods, libname=None):
    libs = doc.BasicLibraries
    names = libs.getElementNames()
    if libname is None:
        libname = 'VBAProject' if 'VBAProject' in names else 'Standard'
    if not libs.hasByName(libname):
        libs.createLibrary(libname)
    lib = libs.getByName(libname)
    for n in list(lib.getElementNames()):
        lib.removeByName(n)
    from com.sun.star.script import ModuleInfo
    for name, src, is_class in mods:
        try:
            if lib.hasModuleInfo(name): lib.removeModuleInfo(name)
            mi = ModuleInfo(); mi.ModuleType = (2 if is_class else 1)  # CLASS=2? NORMAL=1
            lib.insertModuleInfo(name, mi)
        except Exception as e:
            pass
        lib.insertByName(name, src)
    return libname

def run(doc, lib, mod, fn, args=()):
    sp = doc.getScriptProvider()
    s = sp.getScript(f'vnd.sun.star.script:{lib}.{mod}.{fn}?language=Basic&location=document')
    r = s.invoke(tuple(args), (), ())
    return r[0]
