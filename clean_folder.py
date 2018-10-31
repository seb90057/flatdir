
import os
import tools
from filekind.MsgFile import MsgFile
import datetime
from filekind.ZipFile import ZipFile

folder_path = 'C:/Users/sebde/PycharmProjects/project_mgt/data'



path_to_treat = tools.path_to_treat(folder_path)

print(path_to_treat)

while len(path_to_treat) > 0:
    for p in path_to_treat:
        ext = tools.get_extension(p).lower()
        print(ext)
        if ext in ['msg', 'zip']:
            if ext == 'msg':
                msg = MsgFile(p)
                msg.save_attachments()
                del msg
                tools.rename(p, '{}/{}/{}'.format(os.path.dirname(p),
                                               'processed',
                                               tools.add_suffix(os.path.basename(p), datetime.datetime.now().strftime("%Y%m%d%H%M%S"))))
            if ext == 'zip':
                zip = ZipFile(p, timestamp=datetime.datetime.strptime(tools.get_ts_suffix(p), "%Y%m%d%H%M%S"))
                zip.extract_all()
                del zip
                tools.rename(p, '{}/{}/{}'.format(os.path.dirname(p),
                                               'processed',
                                               tools.add_suffix(os.path.basename(p), datetime.datetime.now().strftime("%Y%m%d%H%M%S"))))
        else:
            if tools.get_ts_suffix(p) is None:
                tools.rename(p, '{}/{}'.format(os.path.dirname(p),
                                               tools.add_suffix(os.path.basename(p),
                                                                datetime.datetime.fromtimestamp(os.path.getctime(p)).strftime('%Y%m%d%H%M%S'))))

    path_to_treat = tools.path_to_treat(folder_path)