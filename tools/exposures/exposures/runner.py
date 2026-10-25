#built in
import collections
import json
import logging
import os
import re

#third party
import pyodbc


LOGGER = logging.getLogger(__name__)
TABLE_MAPPING = { 'inforce':{'name':'RawCessions', 'proc':'Stage.UpdateCarrierIDinRawCessions {company}'},
                  'inforce2':{'name':'RawCessions', 'proc':'Stage.UpdateCarrierIDinRawCessions {company}'},
                  'billing':{'name':'RawBilling', 'proc':'Stage.UpdateCarrierIDinRawBilling {company}'},
                  'billing2':{'name':'RawBilling', 'proc':'Stage.UpdateCarrierIDinRawBilling {company}'},
                  'm81pols':{'name':'SLDM81', 'proc':''}, #empty string for no proc to run
                  'updates':{'name':'Updates', 'proc':''},
                  'addlmag':{'name':'AdditionalMagData', 'proc':''},
                  'billret':{'name':'RawBillingRetro', 'proc':''}}

Instruction = collections.namedtuple('Instruction', ['command', 'arguments'])


class QuarterRunner:
    def __init__(self, working_dir, fmt_file_dir, data_dir, year, quarter):
        LOGGER.debug("Creating runner class instance...")
        self.root = working_dir
        self.instructions_path = os.path.join(working_dir, 'instructions.json')
        self.data_dir = data_dir
        self.format_file_dir = fmt_file_dir
        self.instructions = self.get_instructions()
        self.loaded = []
        self.failed_loads = []
        self.year = year
        self.quarter = quarter

    def get_instructions(self, path=None):
        LOGGER.debug("Getting instructions from {}".format(path))
        if not path:
            path = self.instructions_path
        with open(path, 'r') as f:
            instruction_obj = json.load(f, object_pairs_hook=collections.OrderedDict)
        LOGGER.debug(instruction_obj)

        instructions = {}
        for set_name in instruction_obj["instructions"]:
            instructions[set_name] = []
            for i in instruction_obj["instructions"][set_name]:
                LOGGER.debug("Adding instruction: {}".format(i))
                ins = Instruction(*i.split(":"))
                instructions[set_name].append(ins)

        LOGGER.debug("Finished adding instructions")
        LOGGER.debug(instructions)
        return instructions

    def run(self, db, instruction_set=None):
        if not instruction_set:
            instruction_set = self.instructions.keys()
        elif isinstance(instruction_set, str):
            instruction_set = [instruction_set]
        LOGGER.debug(instruction_set)
        for set_name in instruction_set:
            for i in self.instructions[set_name]:
                LOGGER.debug("RUNNER AT: {}".format(i))
                if i.command.lower() == "load":
                    self._load(db, i.arguments)
                elif i.command.lower() == "proc":
                    self._procedure(db, i.arguments)
                elif i.command.lower() == "script":
                    self._script(db, i.arguments)
                else:
                    raise RuntimeError("Unrecognized instruction: "+i.command)
            LOGGER.info('Finished instruction set {}'.format(set_name))

    def _load(self, db, regex):
        LOGGER.info('Load dir {} with regex {}'.format(self.data_dir, regex))
        filelist = []
        for root, _dir, files in os.walk(self.data_dir):
            filelist += [os.path.join(root, f) for f in files]
        txtlist = [f for f in filelist]
        matches = [os.path.join(self.data_dir, c) for c in txtlist if re.search(regex, c)] # can't use abspath since relative to cwd
        notloaded = [m for m in matches if m not in self.loaded]
        batch_size = len(notloaded)
        LOGGER.info("batch size: {}".format(batch_size))
        LOGGER.debug(notloaded)
        error_count = 0
        failures = 0
        for n in notloaded:
            load_settings = self._get_load_settings(n)
            next_id = None
            try:
                tbl = load_settings['table']
                res = db.proc('dbo.LogFiles', [n])
                next_id = res[0][0][0]
                db.bcp(load_settings['path'],
                       load_settings['schema'],
                       load_settings['table'],
                       load_settings['format_file'],
                       quiet=True)
                db.script(f'update stage.{tbl} set fileId = {next_id} where fileId is null')
                if load_settings['proc']:
                    db.proc(load_settings['proc'])
                with open(os.path.dirname(os.path.dirname(os.path.realpath(__file__)))+'\error.txt') as f: #Move to BCP Function
                    file_error_count = int(len([0 for _ in f])/2)#Dropped_Records
                    error_count += file_error_count
                    LOGGER.debug('Loading: {}'.format(n)+'--'+str(file_error_count)+' Dropped Record(s)')
                self.loaded.append(n)
            except Exception as e:
                if next_id is not None: # explicit comparison to None since this could in theory be zero
                    db.script(f'delete stage.{tbl} where fileid = {next_id}')
                    db.script(f'delete dbo.Files where fileid = {next_id}')
                failures += 1
                LOGGER.exception(e)
        log_message = 'load complete: {} files failed, {} dropped records'
        LOGGER.info(log_message.format(failures, error_count))

    def _get_load_settings(self, name):
        """
            Expects absolute paths
        """
        if int(self.year) >= 2022 and int(self.quarter)>=4:
            end_of_name = name.split("_")[-2:]
            company_and_file_type = end_of_name[0]+os.path.splitext(end_of_name[1])[0]
            company = end_of_name[0].upper() 
            file_type = os.path.splitext(end_of_name[1])[0].lower()  
        else:
            end_of_name = name.split("_")[-1]
            company_and_file_type = os.path.splitext(end_of_name)[0]
            company = company_and_file_type[:-7].upper()
            file_type = company_and_file_type[-7:].lower()
        path = os.path.abspath(name)
        fmt_path = os.path.join(self.format_file_dir, company_and_file_type + '.fmt')
        settings = {'table': TABLE_MAPPING[file_type]['name'],
                    'path': path,
                    'proc': TABLE_MAPPING[file_type]['proc'].format(company=company),
                    'company_code': company,
                    'format_file': fmt_path,
                    'schema': 'Stage',
                    'quiet': True}
    
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


