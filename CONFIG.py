from typing import Final

TOKEN: Final[str] = ""

HEADERS: Final[dict] = {'Authorization': TOKEN}


class MyError(Exception):
    def __init__(self, *args):
        if args:
            self.message = args[0]
        else:
            self.message = None

    def __str__(self):
        if self.message:
            return 'MyError, {0} '.format(self.message)
        else:
            return 'MyError has been raised'
