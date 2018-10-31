
import win32com.client
import os
import tools

class MsgFile():

    def __init__(self, path):

        # create outlook object
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.msg = self.outlook.OpenSharedItem(path)

        # original folder
        self.folder = os.path.dirname(path)

        # get msg date
        self.timestamp = self.msg.SentOn

        # get msg attachments
        self.atts = self.msg.Attachments

    def get_timestamp(self):
        return self.timestamp

    def get_attachments(self):

        att_list = []
        count_attachments = self.atts.Count

        if count_attachments > 0:
            for item in range(count_attachments):
                att_list.append(self.atts.Item(item + 1))

        return att_list

    def save_attachments(self, dest_path=None):
        for att in self.atts:

            if dest_path is None:
                dest_path = self.folder

            new_path = '{}/{}'.format(dest_path, tools.add_suffix(att.Filename, tools.timestamp_to_string(self.timestamp)))

            att.SaveAsFile(new_path)

    def __str__(self):

        s = 'timestamp : {}'.format(msg.get_timestamp().strftime('%Y%m%d%H%M%S')) + '\n'
        s += 'nb of attachments : {}'.format(len(msg.get_attachments()))

        return s


if __name__ == '__main__':
    path = 'C:/Users/sebde/PycharmProjects/project_mgt/data/test2.msg'
    msg = MsgFile(path)
    msg.save_attachments()

