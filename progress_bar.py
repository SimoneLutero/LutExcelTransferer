from colored_message import ColoredMessage
import math

class ProgressBar:
  def __init__(self, max_i):
    self.max_i = max_i

  def progress(self, i):
    progress = math.ceil(100*i/self.max_i)
    remaining = 100 - progress
    self.progress_bar = f"[{ColoredMessage.success(progress*'#')}{ColoredMessage.processing(remaining*'-')}] {progress}%"

  def print(self):
    print(self.progress_bar, end='\r', sep='')