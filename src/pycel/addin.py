"""A simple Excel addin, requires www.pyxll.com"""
from pyxll import xl_func, xl_menu, get_config, xl_macro, get_active_object
import win32api
import webbrowser
import os
import win32com.client
from pycel.excelwrapper import ExcelComWrapper
from pycel.excelcompiler import ExcelCompiler


@xl_menu("Open log file", menu="PyXLL")
def on_open_logfile():
    # the PyXLL config is accessed as a ConfigParser.ConfigParser object
    config = get_config()
    if config.has_option("LOG", "path") and config.has_option("LOG", "file"):
        path = os.path.join(config.get("LOG", "path"), config.get("LOG", "file"))
        webbrowser.open("file://%s" % path)


def xl_app():
    xl_window = get_active_object()
    xl_app = win32com.client.Dispatch(xl_window).Application
    return xl_app


@xl_menu("Compile selection", menu="Pycel")
def compile_selection_menu():
    curfile = xl_app().ActiveWorkbook.FullName
    newfile = curfile + ".pickle"
    selection = xl_app().Selection
    seed = selection.Address
    
    if not selection or seed.find(',') > 0:
        win32api.MessageBox(0, "You must select a cell or a rectangular range of cells", "Pycel")
        return
    
    res = win32api.MessageBox(0, "Going to compile %s to %s starting from %s" % (curfile,newfile,seed), "Pycel", 1)
    if res == 2:
        return
    
    sp = do_compilation(curfile, seed)
    win32api.MessageBox(0, "Compilation done, graph has %s nodes and %s edges" % (len(sp.graph.nodes()),
                                                                                  len(sp.graph.edges())),
                        "Pycel")


def do_compilation(filename, seed, sheet=None):
    excel = ExcelComWrapper(filename, app=xl_app())
    c = ExcelCompiler(filename=filename, excel=excel)
    sp = c.gen_graph(seed, sheet=sheet)
    sp.save_to_file(filename + ".pickle")
    sp.export_to_gexf(filename + ".gexf")
    return sp
