
class TaskBean(object):

    __m_project = None
    __m_issuetype = None
    __m_summary = None
    __m_teamdevelopment = None
    __m_Functionality = None
    __m_assignee = None
    __m_teamtest = None
    __m_bugfoundby = None
    __m_labels = None
    __m_description = None
    __m_reproducibility = None
    __m_Severity = None
    __m_Manufacturervariant = None
    __m_foundinSW = None
    __m_foundinHW = None
    __m_key = None
    __m_comments = None

    def getcomments(self):
        return self.__m_comments
    def setcomments(self, __m_comments):
        self.__m_comments = __m_comments

    def getreproducibility(self):
        return self.__m_reproducibility
    def setreproducibility(self, __m_reproducibility):
        self.__m_reproducibility = __m_reproducibility

    def getkey(self):
        return self.__m_key

    def setkey(self, __m_key):
        self.__m_key = __m_key

    def getSeverity(self):
        return self.__m_Severity
    def setSeverity(self, __m_Severity):
        self.__m_Severity = __m_Severity

    def getManufacturervariant(self):
        return self.__m_Manufacturervariant
    def setManufacturervariant(self, __m_Manufacturervariant):
        self.__m_Manufacturervariant = __m_Manufacturervariant

    def getfoundinSW(self):
        return self.__m_foundinSW
    def setfoundinSW(self, __m_foundinSW):
        self.__m_foundinSW = __m_foundinSW

    def getfoundinHW(self):
        return self.__m_foundinHW
    def setfoundinHW(self, __m_foundinHW):
        self.__m_foundinHW = __m_foundinHW

    def getProject(self):
        return self.__m_project
    def setProject(self, _project):
        self.__m_project = _project

    def getIssuetype(self):
        return self.__m_issuetype
    def setIssuetype(self, _issuetype):
        self.__m_issuetype = _issuetype

    def getteamdevelopment(self):
        return self.__m_teamdevelopment
    def setteamdevelopment(self, __m_teamdevelopment):
        self.__m_teamdevelopment = __m_teamdevelopment

    def getFunctionality(self):
        return self.__m_Functionality
    def setFunctionality(self, __m_Functionality):
        self.__m_Functionality = __m_Functionality

    def getteamtest(self):
        return self.__m_teamtest
    def setteamtest(self, __m_teamtest):
        self.__m_teamtest = __m_teamtest

    def getbugfoundby(self):
        return self.__m_bugfoundby
    def setbugfoundby(self, __m_bugfoundby):
        self.__m_bugfoundby = __m_bugfoundby

    def getLabels(self):
        return self.__m_labels
    def setLabels(self, _labels):
        self.__m_labels = _labels


    def getSummary(self):
        return self.__m_summary
    def setSummary(self, _summary):
        self.__m_summary = _summary

    def getDescription(self):
        return self.__m_description
    def setDescription(self, _description):
        self.__m_description = _description


    def getAssignee(self):
        return self.__m_assignee
    def setAssignee(self, _assignee):
        self.__m_assignee = _assignee


    def getDataDict(self, _jiarHndl):
        _project = {}
        _project['key'] = self.getProject()

        _issueType = {}
        _issueType['name'] = self.getIssuetype()

        _summary = self.getSummary()

        _teamdevelopment = self.getteamdevelopment()

        _functionality = []
        _functionality = self.getFunctionality()

        _assignee = {}
        _assignee['name'] = self.getAssignee()

        _teamtest = self.getteamtest()

        _bugfoundby = {}
        _bugfoundby['value'] = self.getbugfoundby()

        _description = self.getDescription()

        _reproducibility = {}
        _reproducibility['value'] = self.getreproducibility()

        _severity = {}
        _severity['value'] = self.getSeverity()

        _manufacturervariant = self.getManufacturervariant()

        _foundinSW = self.getfoundinSW()

        _foundinHW = self.getfoundinHW()

        _labels = []
        if self.getLabels() is not None:
            _labels = self.getLabels()

        issue_dict = {
        'project': _project,
        'issuetype': _issueType,
        'summary': _summary,
        'customfield_10305': str(_teamdevelopment),
        'customfield_10835': str(_functionality),
        'customfield_10306': str(_teamtest),
        'customfield_10121': _bugfoundby,
        'customfield_10118': _reproducibility,
        'customfield_10115': _severity,
        'customfield_13702': str(_manufacturervariant),
        'customfield_14001': str(_foundinSW),
        'customfield_10122': _foundinHW,
        'assignee':_assignee
        }
        #'labels':_labels,
        #'description':_description

        if len(_labels) != 0:
            issue_dict['labels'] = _labels
        if _description is not None:
            issue_dict['description'] = str(_description)
        #print(issue_dict)
        return issue_dict