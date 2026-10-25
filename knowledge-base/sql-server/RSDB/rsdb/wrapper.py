# standard library
import collections
import json
import logging
import os
import re
import shutil
# this project
from .runner import QuarterRunner
from .basedatabase import BaseDatabase
from .retriever import Retriever


LOGGER = logging.getLogger(__name__)

class RsdbWrapper:
    def __init__(self, year, quarter, root, server, database, data=None, pause=False, programs=None):
        LOGGER.debug('Creating RsdbWrapper')

        self.year = str(year) if year else ''  # must be a string for inclusion in path
        self.quarter = str(quarter) if quarter else ''
        
        self.root = root
        self.working = os.path.join(root, 'etl', self.year, 'Q' + self.quarter)
        self.data = data or os.path.join(self.root, 'data', self.year, 'Q' + self.quarter) #os.path.join(self.working, 'data')
        self.format_files = os.path.join(self.root, 'etl', 'format_files')
        self.programs = programs or root
        self.pause = pause

        self.database = BaseDatabase(server, database)
        try:
            self.runner = QuarterRunner(self.working, self.format_files, self.data, self.programs)
        except FileNotFoundError:
            pass # no instruction file found, swallow the error for now in case this is just a reset

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
        if self.pause:
            input(f'>>> Paused between steps! Press enter to reset database to {backup_path}')
        LOGGER.info(f'Resetting database from {backup_name}')
        with self.database as db: 
            db.reset(backup_path)


    def backup(self, path=None):
        if not path:
            path = self.database.database + '.bak'
        backup_path = os.path.join(self.root, path)
        if os.path.isfile(backup_path):
            raise FileExistsError(f'file already exists at {backup_path}')
        if self.pause:
            input(f'>>> Paused between steps! Press enter to back up database to {backup_path}')
        LOGGER.info(f'creating backup at {backup_path}')
        with self.database as db:
            db.make_backup(backup_path)


    def follow_instructions(self, instruction_set):
        if not self.runner:
            raise FileNotFoundError('instruction file not found')
        with self.database as db:
            LOGGER.info('Starting instruction set {}'.format(instruction_set))
            self.runner.run(db, instruction_set, self.pause)
            LOGGER.info('Completed instruction set {}'.format(instruction_set))
    
    def retrieve(self, path):
        assert path is not None, "Retrieve path must be defined"
        golden = Retriever(path, self.data, self.root, self.year, self.quarter)
        LOGGER.info(f'Starting to retrieve files from {path}')
        count = golden.retrieve()
        LOGGER.info(f'Completed retrieving files from {path}, moved {count} files')
        golden.addToDVC()
        LOGGER.info(f'Added files to DVC tracking')