from enum import Enum

class Colors(Enum):
  HEADER = '\033[95m'
  OKBLUE = '\033[94m'
  OKCYAN = '\033[96m'
  OKGREEN = '\033[92m'
  WARNING = '\033[93m'
  FAIL = '\033[91m'
  ENDC = '\033[0m'
  BOLD = '\033[1m'
  UNDERLINE = '\033[4m'

class ColoredMessage:
  def colored_message(color: Colors, message: str):
    return f'{color.value}{message}{Colors.ENDC.value}'

  def success(message):
    return ColoredMessage.colored_message(Colors.OKGREEN, message)

  def processing(message):
    return ColoredMessage.colored_message(Colors.OKCYAN, message)

  def error(message):
    return ColoredMessage.colored_message(Colors.FAIL, message)

  def warning(message):
    return ColoredMessage.colored_message(Colors.WARNING, message)
