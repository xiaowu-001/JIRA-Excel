import wx
import os
import sys
from src.excelLib.CDoExcel import CDoExcel
from src.const._const import _const
from src.jiraLib.CDoJIRA import CDoJIRA

class CMainUIBase(wx.Panel):

    _const.BORDER = 3
    _const.BTN_EXPORT = ''

    _m_JiraUsername = None
    _m_JiraPasswd = None
    _m_pathExcel = None
    _m_jqlText = None
    _m_logText = None
    _m_DoExcel = None
    _m_DoJIRA = None

    _m_btnExport = None

    def __init__(self,parent):
        wx.Panel.__init__(self, parent)        
        self._initUIFGS()
        self._2ndInit()

    def _2ndInit(self):
        pass

    def _initUIFGS(self):
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        fgs = wx.FlexGridSizer(6, 2, 10,10)

        jira_username = wx.StaticText(self,label='JIRA Username: ',style=wx.ALIGN_RIGHT)
        self._m_JiraUsername = wx.TextCtrl(self)
        jira_passwd = wx.StaticText(self,label='JIRA Password: ')
        self._m_JiraPasswd = wx.TextCtrl(self, style=wx.TE_PASSWORD)

        jql_label = wx.StaticText(self,label='JQL: ',style=wx.ALIGN_RIGHT)
        self._m_jqlText = wx.TextCtrl(self)
        path_excel_label = wx.StaticText(self,label='Excel Path: ')
        self._m_pathExcel = wx.TextCtrl(self,style=wx.TE_READONLY)

        log_label = wx.StaticText(self,label='Log: ')
        #self._m_logText = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.HSCROLL | wx.TE_READONLY)
        self._m_logText = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.HSCROLL)

        open_button = wx.Button(self,label = "Open Excel")
        open_button.Bind(wx.EVT_BUTTON,self._openDialog)
        self._m_btnExport = wx.Button(self)

        #title = wx.StaticText(self, label = "Title") 
        #author = wx.StaticText(self, label = "Name of the Author") 
        #review = wx.StaticText(self, label = "Review")
        #tc1 = wx.TextCtrl(self) 
        #tc2 = wx.TextCtrl(self) 
        #tc3 = wx.TextCtrl(self, style = wx.TE_MULTILINE)
        fgs.AddMany([(jira_username), (self._m_JiraUsername, 1, wx.EXPAND), 
                     (jira_passwd), (self._m_JiraPasswd, 1, wx.EXPAND), 
                     (jql_label), (self._m_jqlText, 1, wx.EXPAND), 
                     (path_excel_label), (self._m_pathExcel, 1, wx.EXPAND), 
                     (log_label), (self._m_logText, 1, wx.EXPAND), 
                     (open_button), (self._m_btnExport)])  
        fgs.AddGrowableRow(4, 1)
        fgs.AddGrowableCol(1, 1)  
        hbox.Add(fgs, proportion = 2, flag = wx.ALL|wx.EXPAND, border = 10) 
        self.SetSizer(hbox)
            
    def __initUIBoxSizer(self):
        #myStyle = wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_VERTICAL
        myStyleH = wx.ALIGN_CENTRE
        myStyleV = wx.ALIGN_LEFT
        
        #JIRA username and password input
        jira_username = wx.StaticText(self,label='Username: ')
        self._m_JiraUsername = wx.TextCtrl(self)
        box01 = wx.BoxSizer(wx.HORIZONTAL)
        box01.Add(jira_username,0.2,myStyleH,border=_const.BORDER )
        box01.Add(self._m_JiraUsername,1.8,myStyleH,border=_const.BORDER )
        jira_passwd = wx.StaticText(self,label='Password: ')
        self._m_JiraPasswd = wx.TextCtrl(self)
        box01.Add(jira_passwd,0.2,myStyleH,border=_const.BORDER )
        box01.Add(self._m_JiraPasswd,1.8,myStyleH,border=_const.BORDER )

        #jira_passwd = wx.StaticText(self,label='Password: ')
        #self._m_JiraPasswd = wx.TextCtrl(self)
        #box02 = wx.BoxSizer(wx.HORIZONTAL)
        #box02.Add(jira_passwd,0.2,myStyleH,border=_const.BORDER )
        #box02.Add(self._m_JiraPasswd,1.8,myStyleH,border=_const.BORDER )

        #JQL input
        jql_label = wx.StaticText(self,label='JQL: ')
        self._m_jqlText = wx.TextCtrl(self)
        box1 = wx.BoxSizer(wx.HORIZONTAL)
        box1.Add(jql_label,0.2,myStyleH,border=_const.BORDER )
        box1.Add(self._m_jqlText,1.8,myStyleH,border=_const.BORDER )

        #Excel path text
        path_excel_label = wx.StaticText(self,label='Excel Path: ')
        self._m_pathExcel = wx.TextCtrl(self,style=wx.TE_READONLY)
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        box2.Add(path_excel_label,0.3,myStyleH,border=_const.BORDER )
        box2.Add(self._m_pathExcel,1.7,myStyleH,border=_const.BORDER )

        #Button
        open_button = wx.Button(self,label = "Open Excel")
        open_button.Bind(wx.EVT_BUTTON,self._openDialog)
        export_button = wx.Button(self,label = "Export to JIRA")
        box3 = wx.BoxSizer(wx.HORIZONTAL)
        box3.Add(open_button,1,myStyleH,border=_const.BORDER )
        box3.Add(export_button,1,myStyleH,border=_const.BORDER )

        #Log text
        #self._m_logText = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.HSCROLL | wx.TE_READONLY)
        self._m_logText = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.HSCROLL)
        box4 = wx.BoxSizer(wx.HORIZONTAL)
        box4.Add(self._m_logText,1,myStyleH,border=_const.BORDER )

        globalBox = wx.BoxSizer(wx.VERTICAL)
        globalBox.Add(box01,1,myStyleV,border=_const.BORDER )
        #globalBox.Add(box02,1,myStyleV,border=_const.BORDER )
        globalBox.Add(box1,1,myStyleV,border=_const.BORDER )
        globalBox.Add(box2,1,myStyleV,border=_const.BORDER )
        globalBox.Add(box3,1,myStyleV,border=_const.BORDER )
        globalBox.Add(box4,1,myStyleV,border=_const.BORDER )

        self.SetSizer(globalBox)

    def _openDialog(self, event):
        wildcard = 'All files(*.*)|*.*'
        dialog = wx.FileDialog(None,'select',os.getcwd(),'',wildcard,wx.FD_OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            self._m_pathExcel.SetValue(dialog.GetPath())
            self._addLogText(dialog.GetPath())
            self._operateExcel(dialog.GetPath())
            dialog.Destroy

    def _addLogText(self, logContent):
        self._m_logText.SetValue(self._m_logText.GetValue() + '\n' + logContent)

    def _clearLogText(self):
        self._m_logText.SetValue('')

    def _operateExcel(self, _excelPath):
        self._m_DoExcel = CDoExcel(_excelPath)

    def _initJIRA(self):
        self._m_DoJIRA = CDoJIRA()
        self._m_DoJIRA.setUsername(str(self._m_JiraUsername.Value))
        self._m_DoJIRA.setPassword(str(self._m_JiraPasswd.Value))
        # self._m_DoJIRA.setServer('http://cnninvmjira01:8082/')   #JNN JIRA
        self._m_DoJIRA.setServer('https://jira.jnd.joynext.com/')  #JND JIRA
        #self._m_DoJIRA.setServer('https://jira.jnd.joynext.com/')
        self._m_DoJIRA.setJQL(str(self._m_jqlText.Value))
        self._m_DoJIRA.initialize()
