import sys


def log(*args):
    """In log ra stderr để không lẫn vào file kết quả."""
    print(*args, file=sys.stderr, flush=True)
