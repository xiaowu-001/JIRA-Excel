from jira import JIRA
import time
import jira
from pprint import pprint

class CDoJIRA(object):

    _m_JIRA = None
    _m_JiraUsername = None
    _m_JiraPasswd = None
    _m_JiraServer = None
    _m_JiraJQL = None

    _m_namemap = None

    def __init__(self):
        super().__init__()

    def initialize(self):
        if self._m_JiraServer is None:
            print('Server is null')
            return

        if self._m_JiraUsername is None or self._m_JiraUsername is '':
            print('Username is null')
            return

        if self._m_JiraPasswd is None or self._m_JiraPasswd is '':
            print('Password is null')
            return

        self._m_JIRA = JIRA(options = {'server':self._m_JiraServer,'verify':False}, basic_auth=(self._m_JiraUsername, self._m_JiraPasswd))
        projects = self._m_JIRA.projects()        
        #print(projects)

    def setUsername(self, _username):
        self._m_JiraUsername = _username
        #self._m_JiraUsername = 'zhao_x3'

    def setPassword(self, _passwd):
        self._m_JiraPasswd = _passwd
        #self._m_JiraPasswd = '881220ily.'

    def setServer(self, _server):
        self._m_JiraServer = _server

    def setJQL(self, _jql):
        self._m_JiraJQL = _jql

    def doSearchJQL(self, _jql=None):

        #allfields = self._m_JIRA.fields()
        #name_map = {field['name']:field['id'] for field in allfields}
        #pprint(name_map)
        #self._m_namemap = name_map

        if self._m_JIRA is None:
            print('JIRA is not created')
            return None

        #_fields = 'project, issuetype, parent, components, summary, customfield_10547, fixVersions, duedate, timetracking, assignee, customfield_10531'
        # _fields = 'project, issuetype, parent, components, summary, customfield_10547, fixVersions, duedate, timetracking, assignee, customfield_10531, description, labels, key'
        _fields = 'project, issuetype, summary, customfield_10305, customfield_10835, assignee, customfield_10306, customfield_10121, description, customfield_10118, customfield_10115, customfield_13702, customfield_14001, customfield_10122, labels, comment, status'

        if _jql is not None:
            issues = self._m_JIRA.search_issues(_jql, fields=_fields, maxResults=1000)
            return issues
        
        if self._m_JiraJQL is not None:
            issues = self._m_JIRA.search_issues(self._m_JiraJQL, fields=_fields, maxResults=1000)
            return issues

    def getNamemap(self):
        return self._m_namemap
            

    def createIssue(self, _dictIssue):
        #test sub-task
        #_dictIssue = {'project': {'key': 'MSM'}, 'issuetype': {'name': 'Sub-task'}, 'parent': {'key': 'MSM-413'}, 'components': [{'name': 'HMI'}], 'summary': 'Test for online jira import8899', 'customfield_10547': '2020-08-17 00:00:00', 'fixVersions': [{'name': 'C101'}], 'duedate': '2020-08-30 00:00:00', 'timetracking': {'originalEstimate': '10d'}, 'assignee': {'name': 'zhao_x3'}, 'customfield_10531': {'value': 'Test'}, 'description':'test13', 'labels':['FeinSpec']}
        return self._m_JIRA.create_issue(fields=_dictIssue)

    def transition_issue(self, _issue, _status, _fields):
        return self._m_JIRA.transition_issue(_issue, _status, fields=_fields)

    def assign_issue(self, _issue, _assign):
        return self._m_JIRA.assign_issue(_issue, _assign)

    def add_comment(self, _issue, _comments):
        return self._m_JIRA.add_comment(_issue, _comments)



#Jira = CDoJIRA()
#Jira.setUsername('zhang_t3')
#Jira.setPassword('zT12345678')
#Jira.setServer('http://cnninvmjira01:8082/')
#Jira.setJQL('key = msm-63')
#Jira.initialize()

#issues = Jira.doSearchJQL()
#name_map = Jira.getNamemap()
#print(issues)

#for iss in issues:
#    print('========================')
#    print(iss.fields.summary)
#    iss.delete()
#    pass

#de = CDoExcel('C:\\ZT\\Working\\紧急\\WBS\\WBS-OnlineApp-SOP2_V3_FieldsTemplate.xlsx')
#de.setSheetName('Import')
#max_row = de.getMaxrow()
#if max_row is not None:
#    for i in range(2, max_row+1):
#        issue_dict = de.getDataDict(i)
#        print(issue_dict)
#        #new_issue = Jira.createIssue(issue_dict)
#        print(str(new_issue))
#        print('===========================')

#_fixVersions = []
#_fixVersions.append({'name': 'C101'})
#issue_dict = {
#    'project': {'key': 'MSM'},
#    'issuetype': {'name': 'Task'},
#    'components': [{'name': 'Online'}],
#    'summary': '[DataCollection][BigData(HMI)]CR0026',
#    'customfield_10547':'2020-8-17',
#    'fixVersions':_fixVersions,
#    'duedate':'2020-8-30',
#    'timetracking':{'originalEstimate': '10d'}
#        }

#print(issue_dict)

