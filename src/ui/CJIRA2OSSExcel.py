import sys
import wx
import win32api,win32con
from src.ui.CMainUIBase import CMainUIBase
from src.const._const import _const
from src.bean.TaskBean import TaskBean

class CJIRA2OSSExcel(CMainUIBase):

    _const.BTN_JIRA2OSSEXCEL = 'JIRA -> OSSExcel'

    def __init__(self,parent):
        super().__init__(parent)

    def _2ndInit(self):
        self._m_btnExport.Bind(wx.EVT_BUTTON,self._exportBtnHandler)
        self._m_btnExport.SetLabel(_const.BTN_JIRA2OSSEXCEL)

    def _exportBtnHandler(self, event):
        print('JIRA -> OSS Excel button clicked')

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

        self._m_DoExcel.setSheetName('Vulnerability Management')
        # self._m_DoExcel.setSheetName('Filter Vulnerability ')
        search_issues = self._m_DoJIRA.doSearchJQL(self._m_jqlText.GetValue())
        print(len(search_issues))
        self._clearLogText()
        search_List = []
        for iss in search_issues:
            dict1 = {}
            summary = str(iss.fields.summary)
            jira_num = str(iss.key)
            # for field_name in iss.raw['fields']:
            #     print("Field:", field_name, "Value:", iss.raw['fields'][field_name])   #打印所有的field内容
            status = str(iss.fields.status)
            foundinSW = str(iss.fields.customfield_14001)
            ManufacturerVariant = str(iss.fields.customfield_13702)
            # print("status: ", status)
            if 'CVE' in summary:
                index = summary.index('CVE-')
                dict1.setdefault('cveid', summary[index:])
                dict1.setdefault('component', summary[:index])
                comments = ''
                for comment_body in iss.raw['fields']['comment']['comments']:
                    #不累加，只取最后一个comment
                    comments = 'author: ' + str(comment_body['author']['name']) + '\n' + \
                                         'created_time: ' + str(comment_body['created']) + '\n' + \
                                         'updated_time: ' + str(comment_body['updated']) + '\n' + \
                                         'comment: ' + str(comment_body['body']) + '\n'
                if "是否有影响：是" in comments or "是否有影响：有" in comments or "是否有影响：Yes" in comments:
                    dict1.setdefault('influence', 'Y')
                elif "是否有影响：否" in comments or "是否有影响：无" in comments or "无影响" in comments or "是否有影响：NO" in comments:
                    dict1.setdefault('influence', 'N')
                else:
                    dict1.setdefault('influence', 'NA')
                dict1.setdefault('comments', comments)
                dict1.setdefault('jiranum', jira_num)
                dict1.setdefault('status', status)
                if "339511" in foundinSW:
                    foundinSW = "C363.0-RC6 - CNS_37w"
                if "295" in ManufacturerVariant:
                    ManufacturerVariant = "5HG 035 866 (Online_China) - JPCC"
                AffectedSW = "foundinSW: " + foundinSW + '\n' + "ManufacturerVariant: " + ManufacturerVariant
                dict1.setdefault('AffectedSW', AffectedSW)
                search_List.append(dict1)
                print(search_List)
                self._addLogText("导入CVEID: " + dict1.get('cveid'))
        maxRow = self._m_DoExcel.getMaxrow()
        if maxRow is not None:
            for i in range(3, maxRow+1):
                self._m_DoExcel.setJIRA2OSSData(i, search_List)

        self._m_DoExcel.saveExcel()
        self._addLogText('comments导入完成')
        win32api.MessageBox(0, "导出完成", "提醒",win32con.MB_OK)
            

        