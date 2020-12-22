#!/usr/bin/env python
# -*- coding:utf-8 -*-


import argparse
import passlib
import msoffcrypto


"""usage
./unlock.py -d ${PATH_TO_DATA}/locked_data.xlsx -p PASSWORD
"""


def argParse():
    parser = argparse.ArgumentParser(
        add_help=True
    )
    parser.add_argument(
        '-d',
        '--xlsx-dir',
        dest='xlsx_dir',
        type=str,
        required=True
    )
    parser.add_argument(
        '-p',
        '--pass',
        dest='password',
        type=str,
        required=True
    )
    args = parser.parse_args()
    return args


def unlock(xlsx_dir, password):
    file= pathlib.Path(xlsx_dir)
    with file.open(mode='rb') as locked:
        office_file = msoffcrypto.OfficeFile(locked)
        office_file.load_key(password=password)
        unlocked_file = xlsx_dir.split(".xlsx")[0] + "_unlocked.xlsx"
        with open(unlocked_file, mode='wb') as unlocked:
            office_file.decrypt(unlocked)


if __name__ == "__main__":
    args = argParse()
    unlock(
        password=args.password,
        xlsx_dir=args.xlsx_dir
    )
