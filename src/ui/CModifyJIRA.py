import sys
import wx
import win32api,win32con
#from tkinter import messagebox
from src.ui.CMainUIBase import CMainUIBase
from src.const._const import _const

class CModifyJIRA(CMainUIBase):

    _const.BTN_EXCEL2JIRA = 'Modify -> JIRA'
    

    def __init__(self,parent):
        super().__init__(parent)

    def _2ndInit(self):
        self._m_btnExport.Bind(wx.EVT_BUTTON,self._exportBtnHandler)
        self._m_btnExport.SetLabel(_const.BTN_EXCEL2JIRA)

    def _exportBtnHandler(self, event):
        print('Excel -> JIRA button clicked')
        if self._m_DoExcel is None:
            self._addLogText('Excel is not being opened')
            return None

        if self._m_JiraUsername is None or str(self._m_JiraUsername.Value) is '':
            self._addLogText('JIRA username is null')
            return None

        if self._m_JiraPasswd is None or str(self._m_JiraPasswd.Value) is '':
            self._addLogText('JIRA password is null')
            return None

        self._initJIRA()

        if self._m_DoJIRA is None:
            self._addLogText('JIRA is not being initialized')
            return None

        self._m_DoExcel.setSheetName('Import')
        maxRow = self._m_DoExcel.getMaxrow()
        self._clearLogText()
        if maxRow is not None:
            for i in range(2, maxRow+1):
                issue_dict = self._m_DoExcel.getDataDict(i, self._m_DoJIRA)
                print("wu_T2: " + str(issue_dict))
                #self._addLogText('\"' + str(issue_dict['summary']) + '\" is being created' )
                new_issue = self._m_DoJIRA.createIssue(issue_dict)
                print(str(new_issue))
                # self._m_DoJIRA.add_comment(new_issue, 'TEST1.')
                # self._m_DoJIRA.assign_issue(new_issue, 'zhang_e1')
                # issue_change = {
                #     'customfield_10115': {'value': 'block'}
                # }
                # self._m_DoJIRA.transition_issue(new_issue, '241', issue_change)
                self._addLogText('\"' + str(issue_dict['summary']) + '\" is created successfully' )
            win32api.MessageBox(0, "导入完成", "提醒",win32con.MB_OK)