# standard library
import collections
import json
import logging
import os
import re
import subprocess
import sys
import tempfile

# third party
import pyodbc


LOGGER = logging.getLogger(__name__)
FEATURE_FLAGS = {'use bcp': True}


class Instruction:
    valid_commands = ('load', 'proc', 'script', 'note', 'call', 'pause')
    def __init__(self, command, arguments):
        command = command.lower()
        if command not in self.valid_commands:
            raise ValueError(f'invalid instruction type: {command}')
        self.command = command
        self.arguments = arguments

    def __str__(self):
        return self.command + ':' + self.arguments


class QuarterRunner:
    def __init__(self, working_dir, fmt_file_dir, data_dir, programs_dir):
        LOGGER.debug("Creating runner class instance...")
        self.root = working_dir
        self.instructions_path = os.path.join(working_dir, 'instructions.json')
        self.data_dir = data_dir
        self.programs_dir = programs_dir
        self.format_file_dir = fmt_file_dir
        self.instructions = self.get_instructions()
        self.loaded = []
        self.failed_loads = []

    def get_instructions(self, path=None):
        LOGGER.debug("Getting instructions from {}".format(path))
        if not path:
            path = self.instructions_path
        with open(path, 'r') as f:
            instruction_obj = json.load(f, object_pairs_hook=collections.OrderedDict)
        LOGGER.debug(instruction_obj)

        instructions = {}
        invalid_cmds = set()
        for set_name in instruction_obj["instructions"]:
            instructions[set_name] = []
            for line in instruction_obj["instructions"][set_name]:
                LOGGER.debug("Adding instruction: {}".format(line))
                try:
                    split = line.split(":", 1)
                    cmd = split[0]
                    args = split[1] if len(split) > 1 else 'None'
                    ins = Instruction(cmd, args)
                    instructions[set_name].append(ins)
                except ValueError:
                    invalid_cmds.add(cmd)
 
        if invalid_cmds:
            raise ValueError(f'invalid instruction types: {invalid_cmds}')

        LOGGER.debug("Finished adding instructions")
        LOGGER.debug(instructions)
        return instructions

    def run(self, db, instruction_set, pause=False):
        load_group = 0
        for ins in self.instructions[instruction_set]:
            LOGGER.debug("RUNNER AT: {}".format(ins))
            if pause:
                pause = input(f'>>> Paused between steps! Press enter to run {ins}') != 'c'
            if ins.command == "load":
                self._load(db, ins.arguments, load_group)
                load_group += 1
            elif ins.command == "proc":
                self._procedure(db, ins.arguments)
            elif ins.command == "script":
                self._script(db, ins.arguments)
            elif ins.command == 'call':
                self._call(ins.arguments)
            elif ins.command == "note":
                pass
            elif ins.command == 'pause':
                input(f'>>> Paused between steps! Press enter to run next step. Pause message:\n{ins.arguments}')
            else:
                raise RuntimeError("Unrecognized instruction: " + ins.command)

    def _load(self, db, regex, groupid):
        LOGGER.info('Load dir {} with regex {}'.format(self.data_dir, regex))
        filelist = []
        for root, _dir, files in os.walk(self.data_dir):
            filelist += [os.path.join(root, f) for f in files]
        matches = [os.path.join(self.data_dir, f) for f in filelist if re.search(regex, f)] # can't use abspath since relative to cwd
        notloaded = [m for m in matches if m not in self.loaded]
        batch_size = len(notloaded)
        LOGGER.info("Batch size: {}".format(batch_size))
        LOGGER.debug(notloaded)
        
        failures = []
        unmapped = []
        with tempfile.TemporaryDirectory() as tmpdir:
            error_file = os.path.join(tmpdir, 'error.txt')
            for n in notloaded:
                LOGGER.debug('Loading: {}'.format(n))
                load_settings = self._get_load_settings(n)
                if os.path.isfile(error_file):
                    os.remove(error_file)
                next_id = None
                try:
                    tbl = load_settings['table']
                    res = db.proc('dbo.LogBcpLoad', [n, tbl, groupid])
                    next_id = res[0][0][0] # first result set, first row, first column
                    db.bcp(**load_settings, error_file=error_file)
                    db.script(f'update stage.{tbl} set fileId = {next_id} where fileId is null')
                    self.loaded.append(n)
                except Exception as e:
                    if next_id is not None: # explicit comparison to None since this could in theory be zero
                        db.script(f'delete stage.{tbl} where fileid = {next_id}')
                        db.script(f'delete dbo.DataLoadFiles where loadfileid = {next_id}')
                    LOGGER.error(f'error loading {n}')
                    LOGGER.error(e)
                    failures.append(n)
                    if os.path.isfile(error_file):
                        with open(error_file, 'r') as failfile:
                            LOGGER.debug(failfile.readlines())
        
        error_count = len(failures)
        unmapped_count = len(unmapped)
        success_count = batch_size - error_count - unmapped_count
        log_message = "Load complete: {} successes, {} couldn't be mapped, and {} errors"
        LOGGER.info(log_message.format(success_count, unmapped_count, error_count))

        results = []
        for n in notloaded:
            if n in self.loaded:
                results.append(n+'\tS\n')
            elif n in failures:
                results.append(n+'\tF\n')
            elif n in unmapped:
                results.append(n+'\tU\n')
            else:
                results.append(n+'\t!\n')
        results.append('\n')
        with open('err.txt', 'a') as f:
            f.writelines(results)

    def _get_load_settings(self, name, quiet=True):
        """
            Expects absolute paths
        """
        end_of_name = name.split("_")[-1]
        file_type = os.path.splitext(end_of_name)[0]
        path = os.path.abspath(name)
        fmt_path = os.path.join(self.format_file_dir, file_type + '.fmt')
        settings = {'table': file_type,
                    'path': path, 
                    'format_file': fmt_path,
                    'schema': 'stage',
                    'quiet': quiet}
        return settings

    def _procedure(self, db, procedure):
        LOGGER.info("Procedure: {}".format(procedure))
        if '|' in procedure:
            name, args = procedure.split("|")
            args = args.split(",")
        else:
            name = procedure
            args = []
        db.proc(name, args)

    def _script(self, db, name):
        LOGGER.info("Running script: {}".format(name))
        path = os.path.join(self.root, name)
        with open(path, 'r') as f:
            sql = f.read()
        db.script(sql)

    def _call(self, cmd_str):
        LOGGER.info('calling "{}"'.format(cmd_str))
        cwd = os.getcwd()
        os.chdir(self.programs_dir)
        res = subprocess.run(cmd_str, stderr=subprocess.PIPE, encoding='utf-8', errors='ignore')
        os.chdir(cwd)
        if res.returncode != 0:
            raise RuntimeError(str(res.stderr).replace('\n',' '))
