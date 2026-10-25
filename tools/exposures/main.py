import argparse
import logging
import logging.handlers
import os
import re
import sys


import exposures


LOGGER = logging.getLogger()
PARSER = argparse.ArgumentParser(description="Exposures control program")
PARSER.add_argument('action', nargs='+', help="Method to run, plus year and quarter if needed")
PARSER.add_argument('--root', help="Root directory to find config files, default is current working directory")
PARSER.add_argument('--data', help="Directory to find data files, default is root/YYYY/QQ/data")
PARSER.add_argument('--server', help="Target server, default is 'localhost'")
PARSER.add_argument('--database', help="Target database for insert, default is 'Exposures'")
PARSER.add_argument('--loglevel', default='INFO', help="Minimum level of log output (DEBUG/INFO/WARNING/ERROR/CRITICAL), default is 'INFO'")
PARSER.add_argument('--log_server', nargs=3, help='Two-part log server details, expects "ip:port token"')
PARSER.add_argument('--backup', default=None, help="Path for backup file")


def main():
    args = PARSER.parse_args()

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    LOGGER.addHandler(handler)
    LOGGER.setLevel(args.loglevel.upper())

    action = args.action[0] # should be a single verb for reset or [verb,year,quarter] for everything else
    year = None
    quarter = None
    if action not in ('reset', 'backup'):
        year = args.action[1]
        quarter = args.action[2]

    if args.log_server:
        address, app = args.log_server
        url = "/api/{year}/q/{quarter}/post_log/{app_id}/SHORT/"
        url = url.format(year=year, quarter=quarter, app_id=app)
        httpHandler = logging.handlers.HTTPHandler(address, url)
        LOGGER.addHandler(httpHandler)
        LOGGER.info('Established connection to log server')

    root = args.root or os.getcwd()
    server = args.server or 'localhost'
    database = args.database or 'Exposures'
    data_dir = args.data and re.sub(r'[\'"]$', '', args.data) # remove trailing single and double quotes
    LOGGER.debug('Parsed args')
    wrapper = exposures.ExposuresWrapper(year, quarter, root, server, database, data=data_dir)

    if action == 'create':
        wrapper.create()
    elif action == 'reset':
        wrapper.reset(args.backup)
    elif action == 'backup':
        wrapper.backup(args.backup)
    elif action == 'run':
        wrapper.reset(args.backup)
        wrapper.follow_instructions('refactor')
        wrapper.follow_instructions('load')
        wrapper.follow_instructions('process')
    elif action == 'fastforward':
        wrapper.follow_instructions('refactor')
        wrapper.follow_instructions('load')
        wrapper.follow_instructions('process')
    else:
        wrapper.follow_instructions(action)


if __name__ == '__main__':
    main()
