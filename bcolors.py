
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def colour_print(text, colour):
    if colour == bcolors.OKBLUE:
        string = bcolors.OKBLUE + text + bcolors.ENDC
        print(string)
    elif colour == bcolors.HEADER:
        string = bcolors.HEADER + text + bcolors.ENDC
        print(string)
    elif colour == bcolors.OKCYAN:
        string = bcolors.OKCYAN + text + bcolors.ENDC
        print(string)
    elif colour == bcolors.OKGREEN:
        string = bcolors.OKGREEN + text + bcolors.ENDC
        print(string)
    elif colour == bcolors.WARNING:
        string = bcolors.WARNING + text + bcolors.ENDC
        print(string)
    elif colour == bcolors.FAIL:
        string = bcolors.HEADER + text + bcolors.ENDC
        print(string)
    elif colour == bcolors.BOLD:
        string = bcolors.BOLD + text + bcolors.ENDC
        print(string)
    elif colour == bcolors.UNDERLINE:
        string = bcolors.UNDERLINE + text + bcolors.ENDC
        print(string)
    else:
        assert False, 'Enum color error'

