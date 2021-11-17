from PyQt6.QtCore import *
import os
from typing import List


class CopyFiles(QObject):
    """
    A QObject to copy a list of files to a corresponding list of destinations. The progress is emitted with
    the signal percent_copied and is a value between 0 and 1000 (yes, 1000).
    """
    percent_copied = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self):
        super(CopyFiles, self).__init__()

        self.src = None  # type: List[str]
        self.dest = None  # type: List[str]

        self.auto_start_timer = QTimer()
        self.total_bytes = 0
        self.copy_length = 8 * 1024

    def set_files_to_copy(self, src: List[str], dest: List[str]):
        """
        The files to be copied need to be set with this function before starting the copying.
        :param src:
        :param dest:
        :return:
        """
        self.src = src
        self.dest = dest
        for file in src:
            self.total_bytes += os.path.getsize(file)

    def start_copying(self):
        """
        Where the magic happens.
        """
        copied = 0
        for i in range(len(self.src)):
            if not os.path.exists(os.path.dirname(self.dest[i])):
                os.makedirs(os.path.dirname(self.dest[i]))
            if os.path.exists(self.dest[i]):
                copied += os.path.getsize(self.dest[i])
                self.percent_copied.emit(int(1000 * copied / self.total_bytes))
                continue

            with open(self.src[i], 'rb') as fsrc:
                with open(self.dest[i], 'wb') as fdst:
                    while True:
                        buf = fsrc.read(self.copy_length)
                        if not buf:
                            break
                        fdst.write(buf)
                        copied += len(buf)
                        self.percent_copied.emit(int(1000*copied/self.total_bytes))
        self.finished.emit()
