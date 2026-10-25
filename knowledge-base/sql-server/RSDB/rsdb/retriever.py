import os
import logging
import re
import shutil
import subprocess
import datetime

LOGGER = logging.getLogger(__name__)

class Retriever:
    def __init__(self, path, dest, root, year, quarter, store=True):
        self.path = path # path for pulling dvc files
        self.dest = dest # destination for tracking files
        self.root = root # root directory
        os.makedirs(self.dest, exist_ok=True)

        # make directory to temporarily store previous night's data directory
        self.storageName = f"""AssumedSettlements_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"""
        self.storage = os.path.join(root, self.storageName)
        os.makedirs(self.storage, exist_ok=True)
        self.store = store

        # regex rules for tracking files
        self.rules = [
            "(?!.*(superseded|refactor|alfa))^.*csv$",
            "(?!.*(superseded|refactor))^.*alfa0.txt$",
            "(?!.*(superseded|refactor))^.*alfa1.csv$"
        ]
        self.files = []
        self.count = 0

        LOGGER.debug("Initialized retriever")
        return
    
    def retrieve(self):
        LOGGER.debug(f"Storing prior night's downloads at {self.storage}")
        # move files out of destination directory
        for root, dirs, files in os.walk(self.dest):
            for f in files:
                for rule in self.rules:
                    path = os.path.join(root, f)
                    if re.match(rule, path):
                        shutil.move(path, self.storage)
        
        # zip & delete storage
        if self.store:
            shutil.make_archive(self.storage, 'zip', self.storage)
        shutil.rmtree(self.storage)
        LOGGER.debug(f"Removed storage directory {self.storage}")
        
        # move files over
        for root, dirs, files in os.walk(self.path):
            os.makedirs(os.path.join(self.dest, root[len(self.path)+1:]), exist_ok=True) # creates folder to save in
            for name in files:
                path = os.path.join(root, name)
                for rule in self.rules:
                    if re.match(rule, path):
                        shutil.copy(path, os.path.join(self.dest, root[len(self.path)+1:]))
                        LOGGER.debug(f"Copied {path} to {os.path.join(self.dest, root[len(self.path):])}")
                        self.count += 1        
        return self.count

    def addToDVC(self):
        # change to root directory to run dvc
        wd = os.getcwd()
        os.chdir(self.root)
        
        # iterate through every file, add to dvc
        if self.count > 0:
            for root, dirs, files in os.walk(self.dest):
                for rule in self.rules:
                    for f in files:
                        if re.match(rule, f):
                            LOGGER.info(f"Adding {f} to DVC tracking")
                            subprocess.run(['py', '-m', 'dvc', 'add', os.path.join(root, f)])
                            subprocess.run(['git', 'add', os.path.join(root, f + '.dvc')])
                if os.path.exists(os.path.join(self.dest, root, '.gitignore')):
                    subprocess.run(['git', 'add', os.path.join(self.dest, '.gitignore')])
            subprocess.run(['py', '-m', 'dvc', 'push'])
            
            
            # Janky way of seeing how many .dvc files have been updated, is there something cleaner?
            n = subprocess.run(['git', 'diff', '--cached', '--numstat'], stdout=subprocess.PIPE)
            changed = len(str(n.stdout).split('\\n'))-1
            
            commit_message = f'dvc tracking for {changed} files'
            subprocess.run(['git', 'commit', '-m', commit_message])
            subprocess.run(['git', 'push'])
        os.chdir(wd)