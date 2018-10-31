
import zipfile
import os
import datetime
import tools

class ZipFile():

    def __init__(self, path, timestamp=None):

        # create zip object
        self.zip = zipfile.ZipFile(path)
        self.path = path

        self.timestamp = timestamp

        if timestamp is None:
            self.timestamp = datetime.datetime.fromtimestamp(os.path.getctime(path))

    def extract_all(self, dest=None, add_ts=True):
        if dest is None:
            dest = os.path.dirname(self.path)

        temp_folder = '{}/{}'.format(dest, tools.timestamp_to_string(self.timestamp))

        if not os.path.isdir(temp_folder):
            os.mkdir(temp_folder)

        self.zip.extractall(temp_folder)

        if add_ts:
            file_list = os.listdir(temp_folder)
            for f in file_list:
                tools.rename('{}/{}'.format(temp_folder, f), '{}/{}'.format(dest, tools.add_suffix(f, tools.timestamp_to_string(self.timestamp))))
            os.rmdir(temp_folder)


if __name__ == '__main__':
    path = 'C:/Users/sebde/PycharmProjects/project_mgt/data/test_20181011005436.zip'
    z = ZipFile(path)
    z.extract_all()

