import win32com.client
import os
import zipfile
import datetime
import re

def get_extension(path):
    return path.split('.')[-1]


def extract_attachement(path, target_path=None, rename=True):

    if target_path is None:
        target_path = '/'.join(path.split('/')[0:-1])

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(path)

    timestamp = msg.SentOn.strftime("%Y%m%d%H%M%S")

    # os.mkdir('{}/{}'.format(target_path, timestamp))
    files = []
    count_attachments = msg.Attachments.Count
    att = None
    if count_attachments > 0:
        for item in range(count_attachments):
            att = msg.Attachments.Item(item + 1)
            new_file_name = '{}_{}.{}'.format('.'.join(att.Filename.split('.')[0:-1]), timestamp, get_extension(att.Filename))
            att.SaveAsFile('{}/{}'.format(target_path, new_file_name))
            files.append('{}/{}'.format(target_path, new_file_name))

    del outlook, msg

    return files


def unzip_file(path):

    zip = zipfile.ZipFile(path)
    zip.extractall(path=os.path.dirname(path))

    return os.path.dirname(path)

def get_last_mod(path):
    lm = os.path.getctime(path)
    str = datetime.datetime.fromtimestamp(lm).strftime('%Y%m%d%H%M%S')

    print(str)

def treat_folder(path):
    o_list = os.listdir(path)
    os.mkdir('{}/{}'.format(path, datetime.datetime.now().strftime('%Y%m%d%H%M%S')))

    for o in o_list:
        if os.path.isfile(o):
            if get_extension(o) == "msg":
                extract_attachement(o)


def add_suffix(file_name, suffix):

    return '{}_{}.{}'.format('.'.join(file_name.split('.')[0:-1]), suffix, get_extension(file_name))


def timestamp_to_string(ts):

    return ts.strftime('%Y%m%d%H%M%S')

def path_to_treat(folder_path):

    file_list = [f for f in os.listdir(folder_path) if os.path.isfile('{}/{}'.format(folder_path, f))]

    path_list = ['{}/{}'.format(folder_path, f) for f in file_list]

    msg_path = [p for p in path_list if get_extension(p).lower() == 'msg']
    zip_path = [p for p in path_list if get_extension(p).lower() == 'zip']
    no_ts_path = [p for p in path_list if get_ts_suffix(p) is None]

    return list(set(msg_path + zip_path + no_ts_path))


def get_ts_suffix(path):
    ext = get_extension(path)
    pattern = '^.*(?P<ts>[0-9]{{14}})\.{}$'.format(ext)

    if re.match(pattern, path):
        return re.search(pattern, path).group('ts')
    else:
        return None

if __name__ == '__main__':
    print(get_ts_suffix('test_20181011005436.zip'))