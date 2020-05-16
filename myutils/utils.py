import glob
from contextlib import contextmanager
from datetime import datetime
import os
import sys


def create_folder(nest, path, *args):
    """nest: bool, path: str, *args: folder name
    """
    if not nest:
        base = [os.path.dirname(path)]
        base.extend(args)
        folder = '{}_{}'.format(
                os.path.join(*base),
                datetime.now().strftime('%Y%m%d%H%M%S'))
    else:
        base = [path]
        base.extend(args)
        folder = os.path.join(*base)
    os.makedirs(folder)
    return folder


def get_file(folder, file, suffix=None, ext='xlsx'):
    """makes output file name using original file name
    """
    file_name = os.path.splitext(os.path.basename(file))[0]
    if suffix:
        output_file = os.path.join(
            folder,
            '{}_{}.{}'.format(file_name, suffix, ext)
        )
    else:
        output_file = os.path.join(
            folder,
            '{}.{}'.format(file_name, ext)
        )
    return output_file


@contextmanager
def show_progress():
    def _progress(total_num):
        sys.stdout.write('\rTotal number of checked files: {}'.format(total_num))
        sys.stdout.flush()
    try:
        yield _progress
    finally:
        print()


def confirm(msg):
    while True:
        ans = input('{} y/n: '.format(msg))
        if ans in {'y', 'n'}:
            break
    return ans


def recursive_search(path, target):
    for item in glob.iglob('{}/**/*{}'.format(path, target), recursive=True):
        yield item


def walk_search(path, folder_name, file_name, ext='.html'):
    for root, _, files in os.walk(path):
        if os.path.basename(root) == folder_name:
            for file in files:
                if os.path.splitext(file)[1] == ext:
                    yield os.path.join(root, file)
                    
        for file in files:
            if file == file_name:
                yield os.path.join(root, file)

