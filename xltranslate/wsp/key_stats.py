#
import logging

log = logging.getLogger(__name__)


class KeyStats(object):
    def __init__(self, sheet):
        self._sheet = sheet

    def dump(self):
        print("--Nothing here yet--")
