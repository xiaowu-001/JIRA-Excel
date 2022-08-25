from enum import Enum

class ENENT_TYPE(Enum):
    MESSAGEBOX = 0

class MessageHandler(object):
    def __init__():
        super().__init__()


    def onEvent(self, _eventType, _msgTitle='', _msgContent=''):
        if _eventType is ENENT_TYPE.MESSAGEBOX:
            pass

    def handleMessagebox(self, _msgTitle='', _msgContent=''):
        pass