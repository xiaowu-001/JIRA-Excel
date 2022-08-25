import sys
import wx
from src.ui.CJIRA2Excel import CJIRA2Excel
from src.ui.CExcel2JIRA import CExcel2JIRA
from src.ui.CJIRA2OSSExcel import CJIRA2OSSExcel
from src.const._const import _const

if __name__ == '__main__':
    #sys.modules[__name__] = _const()
    app = wx.App(False)
    frame = wx.Frame(None, title="JIRA-Excel Demo")
    frame.SetDimensions(100,100,640,480)
    nb = wx.Notebook(frame)
    nb.AddPage(CJIRA2Excel(nb),"JIRA->Excel")
    nb.AddPage(CExcel2JIRA(nb),"Excel->JIRA")
    nb.AddPage(CJIRA2OSSExcel(nb), "JIRA->OSSExcel")
    frame.Show()
    app.MainLoop()