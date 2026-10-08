# Opens the source workbook in LibreOffice, freezes panes on the first two sheets and saves it as .ods.
# Run with LibreOffice's own Python: python make-libreoffice-ods.py <source.xlsx> <out.ods>
# Needs a window (invisible): with --headless LibreOffice saves no view settings, so no frozen panes.
import pathlib, subprocess, sys, tempfile, time, uno
from com.sun.star.beans import PropertyValue
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
PROFILE = (pathlib.Path(tempfile.gettempdir()) / "xlsxflow-lo-profile").as_uri()
def prop(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
src, dst = sys.argv[1:3]
office = subprocess.Popen([SOFFICE, "--invisible", "--norestore", "--nologo", "--nodefault", f"-env:UserInstallation={PROFILE}", "--accept=socket,host=127.0.0.1,port=2084;urp;"])
try:
    resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", uno.getComponentContext())
    for _ in range(60):
        try:
            ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=2084;urp;StarOffice.ComponentContext"); break
        except Exception:
            time.sleep(1)
    time.sleep(3)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, ())
    c = doc.getCurrentController()
    sheets = doc.Sheets
    c.setActiveSheet(sheets.getByIndex(0)); c.freezeAtPosition(1, 2)
    if sheets.Count > 1:
        c.setActiveSheet(sheets.getByIndex(1)); c.freezeAtPosition(0, 1)
    c.setActiveSheet(sheets.getByIndex(0))
    doc.storeToURL(uno.systemPathToFileUrl(dst), (prop("FilterName", "calc8"),))
    doc.close(True)
    print("ok")
finally:
    try: desktop.terminate()
    except Exception: pass
    office.wait(timeout=30)
