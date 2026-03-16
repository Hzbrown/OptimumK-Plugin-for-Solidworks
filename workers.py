import sys
from PyQt5.QtCore import QThread, pyqtSignal, QObject
from solidworks_release import release_solidworks_command_state


class QtStream(QObject):
    """Redirect stdout to a PyQt signal."""
    text_written = pyqtSignal(str)

    def write(self, text):
        if text.strip():
            self.text_written.emit(text.strip())

    def flush(self):
        pass


class WorkerBase(QThread):
    """Base class for SolidWorks subprocess worker threads.

    Subclasses set STATE_DESCRIPTIONS and SUCCESS_MESSAGE as class variables
    to customise per-operation state labels and the completion message.
    """
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(int, int)   # current, total
    state_changed = pyqtSignal(str)   # human-readable state description
    log = pyqtSignal(str)

    STATE_DESCRIPTIONS: dict = {}
    SUCCESS_MESSAGE: str = "Operation completed successfully"

    def __init__(self, operation, *args):
        super().__init__()
        self.operation = operation
        self.args = args
        self._abort = False
        self._process = None
        self._total_tasks = 0
        self._current_progress = 0

    def abort(self):
        """Request abort and terminate any running subprocess."""
        self._abort = True
        if self._process and self._process.poll() is None:
            try:
                self._process.terminate()
                self._process.wait(timeout=2)
            except:
                try:
                    self._process.kill()
                except:
                    pass

        if self._process is not None:
            released, release_message = release_solidworks_command_state()
            if not released:
                print(f"Warning: Failed to release SolidWorks state after abort: {release_message}")

    def parse_output_line(self, line):
        """Parse TOTAL:/PROGRESS:/STATE: protocol lines; return None to suppress from log."""
        if line.startswith("TOTAL:"):
            try:
                self._total_tasks = int(line.split(":")[1])
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None
        elif line.startswith("PROGRESS:"):
            try:
                self._current_progress = int(line.split(":")[1])
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None
        elif line.startswith("STATE:"):
            state = line.split(":")[1]
            description = self.STATE_DESCRIPTIONS.get(state, state)
            self.state_changed.emit(description)
            return None
        return line

    def run(self):
        stream = QtStream()
        stream.text_written.connect(self._handle_log)
        old_stdout = sys.stdout
        sys.stdout = stream
        try:
            self.operation(*self.args, worker=self)
            if self._abort:
                self.finished.emit(False, "Operation aborted")
            else:
                self.finished.emit(True, self.SUCCESS_MESSAGE)
        except Exception as e:
            if self._abort:
                self.finished.emit(False, "Operation aborted")
            else:
                self.finished.emit(False, str(e))
        finally:
            sys.stdout = old_stdout

    def _handle_log(self, text):
        result = self.parse_output_line(text)
        if result is not None:
            self.log.emit(result)
