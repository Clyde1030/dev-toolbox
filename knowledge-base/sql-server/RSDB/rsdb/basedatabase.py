from functools import wraps
import itertools
import logging
import os
import re
import subprocess
import time


import pyodbc


LOGGER = logging.getLogger(__name__)
DEFAULT_DRIVER = '{SQL Server Native Client 11.0}'


def _connect_if_needed(func):
    @wraps(func)
    def wrapped_func(self, *args, **kwargs):
        if not self.conn:
            self._connect()
            new_connection = True
        else:
            new_connection = False

        res = func(self, *args, **kwargs)

        if new_connection and not self.keep_alive:
            self.commit()
            self._close()
        
        return res
    return wrapped_func


class BaseDatabase:
    def __init__(self, server, database, driver=DEFAULT_DRIVER):
        LOGGER.debug('Creating BaseDatabase instance...')
        self.server = server
        self.database = database
        self.driver = driver
        self.keep_alive = False
        self.connection_id = None
        self.conn = None
        self.cursor = None

    def __enter__(self):
        LOGGER.debug('Entering BaseDatabase ({})'.format(self.database))
        self.keep_alive = True
        return self
    
    def __exit__(self, type, value, traceback):
        LOGGER.debug('Exiting BaseDatabase ({})'.format(self.database))
        self.keep_alive = False
        if self.conn:
            if type:
                self.conn.rollback()
            else:
                self.conn.commit()
        self._close()

    def _kill(self, *, connection_id=None, all=False):
        self._connect('master', autocommit=True)
        if all:
            sql = """
                DECLARE @kill varchar(8000) = '';  
                IF OBJECT_ID('master..sysprocesses') IS NULL
                BEGIN
                    SELECT @kill = @kill + 'KILL ' + CONVERT(varchar(5), session_id) + ';'  
                    FROM sys.dm_exec_sessions
                    WHERE database_id  = db_id('{db}') and session_id <> {spid}
                END
                ELSE
                BEGIN
                    SELECT @kill = @kill + 'KILL ' + CONVERT(varchar(5), spid) + ';'  
                    FROM master..sysprocesses
                    WHERE dbid = db_id('{db}') and spid <> {spid}
                END
                EXEC (@kill)
            """.format(db=self.database, spid=self.connection_id)
        else:
            sql = 'KILL {}'.format(connection_id)
        LOGGER.debug(sql)
        self.cursor.execute(sql)

    def _connect(self, target_db=None, autocommit=True):
        db = target_db or self.database
        LOGGER.debug('Connecting to {}...'.format(db))
        connection_string = (r'DRIVER=' + self.driver + r';'
                             r'SERVER=' + self.server + r';'
                             r'DATABASE=' + db + r';'
                             r'Trusted_Connection=yes')
        LOGGER.debug('Connection string: {}'.format(connection_string))
        self.conn = pyodbc.connect(connection_string, autocommit=autocommit)
        self.cursor = self.conn.cursor()
        self.cursor.fast_executemany = True
        self.cursor.execute('SET NOCOUNT ON')
        self.connection_id = self.cursor.execute('select @@SPID as spid').fetchone().spid
        LOGGER.debug('spid is {}'.format(self.connection_id))

    def _close(self):
        LOGGER.debug('Closing connection...')
        if self.conn:
            self.conn.close()
        self.conn = None
        self.cursor = None
        self.connection_id = None

    def reset(self, backup_path, auto_kill=True):
        if self.conn:
            raise RuntimeError("Reset failed, must close connection before restoring database {}.".format(self.database))
        
        LOGGER.debug('Resetting...')
        if auto_kill:
            self._kill(all=True)

        # get logical file names
        abs_backup_path = os.path.abspath(backup_path)
        self._connect('master', autocommit=True)
        self.cursor.execute("EXEC('RESTORE FILELISTONLY FROM DISK = N''' + ? + '''')", abs_backup_path) # using exec to build string and force NVARCHAR for argument. there is probably a better way to force NVARCHAR.
        file_list = self.cursor.fetchall()

        # generate physical file names and logical > physical mapping
        self.cursor.execute("""
            SELECT
                CONVERT(VARCHAR(MAX), SERVERPROPERTY('InstanceDefaultDataPath')) as dataPath,
                CONVERT(VARCHAR(MAX), SERVERPROPERTY('InstanceDefaultLogPath')) as logPath
                """) # note the type conversion since return value of SERVERPROPERTY is 'sql-variant' which is not supported by pyodbc
        default_paths = self.cursor.fetchall()[0]
        paths = {'D': default_paths[0] + os.sep + self.database + '_{}_.mdf',
                 'L': default_paths[0] + os.sep + self.database + '_{}_.ldf'}
        mapping = [(f.LogicalName, paths[f.Type].format(i))
                   for i, f in enumerate(file_list)]
        LOGGER.debug(mapping)

        # generate restore query
        restore_query = ('RESTORE DATABASE ? '
                         'FROM DISK = ? '
                         'WITH RECOVERY')
        restore_query += ', MOVE ? TO ?' * len(mapping)
        params = [self.database,
                  abs_backup_path,
                  *itertools.chain.from_iterable(mapping)]

        # Note that the call to time.sleep at the end of this section (specifically, after 
        # cursor.execute) is necessary to prevent a connection error that will hang the newly
        # restored DB and leave it in an unusable state.
        LOGGER.debug(restore_query)
        LOGGER.debug(params)
        LOGGER.debug('Starting restore query...')
        self.cursor.execute(restore_query, *params)
        LOGGER.debug('Restore complete, waiting...')
        time.sleep(5)
        self._close()


    def make_backup(self, path, name='ReinsuranceSettlements-BackupFromPython', auto_kill=True):
        if os.path.isfile(path):
            raise FileExistsError(f'file {path} already exists')
        if self.conn:
            raise RuntimeError(f'backup failed, must close connection before backing up database {self.database}')

        if auto_kill:
            LOGGER.debug('closing connections prior to backup...')
            self._kill(all=True)

        self._connect('master', autocommit=True)
        query = f"BACKUP DATABASE ? TO DISK = ? WITH NOFORMAT, INIT, NAME = ?"
        LOGGER.debug(query)
        LOGGER.debug('backing up database...')
        self.cursor.execute(query, self.database, path, name)
        time.sleep(1)
        LOGGER.debug('backup complete...')
        self._close()


    def bcp(self, path, schema, table, format_file, error_file, quiet=False):
        """ Use bcp to load file from path into schema.table
            format_file and error_file should be paths
        """
        cmd = 'bcp {sch}.{tbl} in "{file}" -f "{fmt}" -S {srvr} -d {db} -T -m 0 -e {efile}' # this is the one place where max_errors=1 implies quit on the first....
        cmd = cmd.format(sch=schema, tbl=table, file=path, fmt=format_file, srvr=self.server, db=self.database, efile=error_file)
        if quiet:
            res = subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        else:
            res = subprocess.run(cmd)
        if res.returncode:
            raise RuntimeError('bcp returned an error: {}'.format(res))
        with open(error_file, 'r') as f:
            file_error_count = int(len([0 for _ in f]) / 2) # Dropped_Records
            if file_error_count > 0:
                raise RuntimeError('bcp dropped {} record(s) from {}'.format(file_error_count, path))


    @_connect_if_needed
    def load(self, data, schema, table, column_mapping=None):
        if not column_mapping:
            column_mapping = ['?'] * len(data[0])
        metadata = self.conn.cursor.columns(schema=schema, table=table).fetchall()
        tbl_cols = [m[3] for m in metadata] # third item is column name
        columns_to_fill = tbl_cols[:len(column_mapping)]
        query = 'insert into {schema}.{table}({cols}) values({vals})'
        query = query.format(schema=schema,
                            table=table,
                            cols=','.join(columns_to_fill),
                            vals=','.join(column_mapping))
        self.cursor.executemany(query, data)


    @_connect_if_needed
    def proc(self, name, values=[]):
        """
            Calls a proc in RPC format, per ODBC spec.
        """
        param_string = ('?, ' * len(values))[0:-2]
        sql = """
            declare @returnVal int
            EXEC @returnVal = {proc} {args}
            select @returnVal
        """.format(proc=name, args=param_string)
        LOGGER.debug(sql)
        self.cursor.execute(sql, values)
        
        results = []
        try:
            results.append(self.cursor.fetchall())
        except pyodbc.ProgrammingError:
            pass
        while self.cursor.nextset():
            try:
                results.append(self.cursor.fetchall())
            except pyodbc.ProgrammingError:
                pass
        return_value = results.pop()[0][0]
        if return_value:
            raise RuntimeError('non-zero return value from database procedure')
        return results
   
    @_connect_if_needed
    def script(self, script_str):
        go_parse_regex = re.compile(r'^\s*[Gg][Oo]\s*$', re.MULTILINE)
        script_with_proper_go = re.sub(go_parse_regex, 'GO', script_str)
        batches = script_with_proper_go.split('GO\n')
        for b in batches:
            trimmed = re.sub(r'^\s*$', '', b)
            if trimmed:
                try:
                    self.cursor.execute(trimmed)
                    if self.cursor.rowcount:
                        try:
                            rows = self.cursor.fetchall()
                        except pyodbc.ProgrammingError:
                            pass
                    while self.cursor.nextset():
                        if self.cursor.rowcount:
                            try:
                                rows = self.cursor.fetchall()
                            except pyodbc.ProgrammingError:
                                pass
                except:
                    LOGGER.debug(trimmed)
                    raise
