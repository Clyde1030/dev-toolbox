# standard library
import argparse
import logging
import logging.handlers
import os
import re
import sys
# this project
import rsdb


LOGGER = logging.getLogger()


def main():
    args = get_args()
    action = args.action[0] # should be a single verb for reset or [verb,year,quarter] for everything else
    year = None
    quarter = None
    if action not in ('reset', 'backup'):
        year = args.action[1]
        quarter = args.action[2]
    configure_logger(LOGGER, args.loglevel, args.log_server, year, quarter)

    root_dir = os.path.abspath(args.root)
    data_dir = args.data and re.sub(r'[\'"]$', '', args.data) # remove trailing single or double quotes (happens when directories are passed in as "J:\whatever\" including the trailing slash)
    programs_dir = (args.programs and re.sub(r'[\'"]$', '', args.programs)) or root_dir
    LOGGER.debug('parsed args')
    wrapper = rsdb.RsdbWrapper(year, quarter, root_dir, args.server, args.database,
                               data=data_dir, pause=args.pause, programs=programs_dir)

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
    elif action == "sprint":
        wrapper.retrieve(args.retrieve)
        wrapper.reset(args.backup)
        wrapper.follow_instructions('refactor')
        wrapper.follow_instructions('load')
        wrapper.follow_instructions('process')
    elif action == 'fastforward':
        wrapper.follow_instructions('refactor')
        wrapper.follow_instructions('load')
        wrapper.follow_instructions('process')
    elif action == 'retrieve':
        wrapper.retrieve(args.retrieve)
    else:
        wrapper.follow_instructions(action)
    LOGGER.info('process complete')


def get_args():
    parser = argparse.ArgumentParser(description="RSDB control program")
    parser.add_argument('action', nargs='+', help="Method to run, plus year and quarter if needed")
    parser.add_argument('--root', default='.', help="Root directory to find config files, default is current working directory")
    parser.add_argument('--data', default=None, help="Directory to find data files, default is root/YYYY/QQ/data")
    parser.add_argument('--server', default='localhost', help="Target server, default is 'localhost'")
    parser.add_argument('--database', default='ReinsuranceSettlements', help="Target database for insert, default is 'ReinsuranceSettlements'")
    parser.add_argument('--loglevel', default='INFO', help="Minimum level of log output (DEBUG/INFO/WARNING/ERROR/CRITICAL), default is 'INFO'")
    parser.add_argument('--log_server', nargs=3, help='Two-part log server details, expects "ip:port token"')
    parser.add_argument('--pause', dest='pause', action='store_true', help='Pause between steps, enter c to play remainder of instruction set without pauses')
    parser.add_argument('--programs', default=None, help="Directory to find external programs, default is to use root")
    parser.add_argument('--backup', default=None, help="Path for backup file")
    parser.add_argument('--retrieve', default=None, help="Path to retrieve files for DVC")
    parser.set_defaults(pause=False)

    args = parser.parse_args()

    args.root = os.path.abspath(args.root)
    if args.data:
        args.data = os.path.abspath(args.data)

    return args


def configure_logger(logger, level, log_server, year, quarter):
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    log_level = level.upper()
    logger.setLevel(log_level)
    if log_server:
        address, app = log_server
        url = "/api/{year}/q/{quarter}/post_log/{app_id}/SHORT/"
        url = url.format(year=year, quarter=quarter, app_id=app)
        httpHandler = logging.handlers.HTTPHandler(address, url)
        logger.addHandler(httpHandler)
        logger.info('Established connection to log server')

if __name__ == '__main__':
    main()
