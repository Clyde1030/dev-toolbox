import collections
import json
import logging
import os
import re
import shutil


from .runner import QuarterRunner
from .basedatabase import BaseDatabase


LOGGER = logging.getLogger(__name__)


class ExposuresWrapper:
    def __init__(self, year, quarter, root, server, database, data=None):
        LOGGER.debug('Creating ExposuresWrapper')

        self.year = str(year) if year else ''  # must be a string for inclusion in path
        self.quarter = str(quarter) if quarter else ''
        
        self.root = root
        self.working = os.path.join(root, 'etl', self.year, self.quarter)
        self.data = data or os.path.join(self.root, 'data', self.year, self.quarter)
        self.format_files = os.path.join(self.root, 'etl', 'format_files')

        self.database = BaseDatabase(server, database)
    
    def create(self):
        LOGGER.info('Creating new filestructure {} {}'.format(self.year, self.quarter))
        LOGGER.debug('Creating {}'.format(self.data))
        os.makedirs(self.data, exist_ok=True)

        for template in ['run_template.json', 'prep_script.sql']:
            dst = os.path.join(self.working, template)
            if not os.path.isfile(dst):
                src = os.path.join(self.root, 'etl', 'templates', template)
                LOGGER.debug('Copying template {} to {}'.format(src, dst))
                shutil.copy2(src, dst)

    def reset(self, backup_name=None):
        if not backup_name:
            backup_name = self.database.database + '.bak'
        backup_path = os.path.join(self.root, backup_name)
        LOGGER.info('Resetting database...')
        with self.database as db: 
            db.reset(backup_path)

    def backup(self, path=None):
        if not path:
            path = self.database.database + '.bak'
        backup_path = os.path.join(self.root, path)
        if os.path.isfile(backup_path):
            raise FileExistsError(f'file already exists at {backup_path}')
        LOGGER.info(f'creating backup at {backup_path}')
        with self.database as db:
            db.make_backup(backup_path)

    def follow_instructions(self, instruction_set):
        LOGGER.info('Starting instruction set {}'.format(instruction_set))
        runner = QuarterRunner(self.working, self.format_files, self.data, self.year, self.quarter)
        if instruction_set not in runner.instructions:
            LOGGER.error('Instruction set "{}" not found'.format(instruction_set))
            return
        with self.database as db:
            runner.run(db, instruction_set)
