import sys
import wx
import win32api,win32con
from src.ui.CMainUIBase import CMainUIBase
from src.const._const import _const
from src.bean.TaskBean import TaskBean

class CJIRA2Excel(CMainUIBase):

    _const.BTN_JIRA2EXCEL = 'JIRA -> Excel'

    def __init__(self,parent):
        super().__init__(parent)

    def _2ndInit(self):
        self._m_btnExport.Bind(wx.EVT_BUTTON,self._exportBtnHandler)
        self._m_btnExport.SetLabel(_const.BTN_JIRA2EXCEL)

    def _exportBtnHandler(self, event):
        print('JIRA -> Excel button clicked')

        if self._m_DoExcel is None:
            self._addLogText('Excel is not being opened')
            return None

        if self._m_JiraUsername is None  or str(self._m_JiraUsername.Value) is '':
            self._addLogText('JIRA username is null')
            return None

        if self._m_JiraPasswd is None or str(self._m_JiraPasswd.Value) is '':
            self._addLogText('JIRA password is null')
            return None

        if self._m_jqlText is None or str(self._m_jqlText.Value) is '':
            self._addLogText('JQL is null')
            return None

        print(self._m_JiraUsername.GetValue())
        print(self._m_JiraPasswd.GetValue())
        print(self._m_jqlText.GetValue())

        self._initJIRA()

        if self._m_DoJIRA is None:
            self._addLogText('JIRA is not being initialized')
            return None

        self._m_DoExcel.setSheetName('Export')
        search_issues = self._m_DoJIRA.doSearchJQL(self._m_jqlText.GetValue())
        print(len(search_issues))
        i = 2
        self._clearLogText()
        for iss in search_issues:
            print('====================')
            # print(iss.fields.project)
            # print(iss.fields.issuetype)
            # print(iss.fields.summary)
            # print(iss.fields.customfield_10305)
            # print(iss.fields.customfield_10835)
            # print(iss.fields.customfield_10306)
            # print(iss.fields.customfield_10118.value)
            # print(iss.fields.customfield_10115.value)
            # print(iss.fields.customfield_13702)
            # print(iss.fields.customfield_14001)
            # print(iss.fields.customfield_10122)
            # print(iss.fields.assignee.name)
            # if len(iss.fields.labels) != 0:
            #     print(iss.fields.labels[0])
            # print(iss.fields.description)
            # for field_name in iss.raw['fields']:
            #     print("Field:", field_name, "Value:", iss.raw['fields'][field_name])   #打印所有的field内容
            # for comment_body in iss.raw['fields']['comment']['comments']:
            #     print(comment_body['author'])
            # if iss.fields.comments is not None:
            #     for comment in iss.comments:
            #         print(comment)
            comments = ''
            for comment_body in iss.raw['fields']['comment']['comments']:
                comments = comments +'author: ' + str(comment_body['author']['name']) + '\n' + \
                                     'created_time: ' + str(comment_body['created']) + '\n' + \
                                     'updated_time: ' + str(comment_body['updated']) + '\n' + \
                                     'comment: ' + str(comment_body['body']) + '\n'
            tb = TaskBean()
            tb.setkey(str(iss.key))
            tb.setProject(str(iss.fields.project))
            tb.setIssuetype(str(iss.fields.issuetype))
            tb.setSummary(str(iss.fields.summary))
            tb.setteamdevelopment(str(iss.fields.customfield_10305))
            tb.setFunctionality(str(iss.fields.customfield_10835))
            tb.setteamtest(str(iss.fields.customfield_10306))
            tb.setreproducibility(str(iss.fields.customfield_10118.value))
            tb.setSeverity(str(iss.fields.customfield_10115.value))
            tb.setManufacturervariant(str(iss.fields.customfield_13702))
            tb.setfoundinSW(str(iss.fields.customfield_14001))
            tb.setfoundinHW(str(iss.fields.customfield_10122))
            tb.setAssignee(str(iss.fields.assignee.name))
            labels = ''
            for label in iss.fields.labels:
                labels = labels + str(label) + ' '
            tb.setLabels(labels)
            tb.setDescription(str(iss.fields.description))
            tb.setcomments(comments)

            self._m_DoExcel.setData2Excel(i, tb)
            self._addLogText('Export \"' + tb.getSummary() + '\"')
            i = i + 1

        self._m_DoExcel.saveExcel()
        win32api.MessageBox(0, "导出完成", "提醒",win32con.MB_OK)
            

        