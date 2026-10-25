@SETLOCAL ENABLEDELAYEDEXPANSION & python -x "%~f0" %* & (IF ERRORLEVEL 1 PAUSE) & EXIT /B !ERRORLEVEL!
"""
    This is a quick-and-dirty script to run the validation proc on files that are dragged and dropped.
    Hence, it is designed to work without installing any dependencies. Hence, it has some workarounds.
    ... and minimal error handling!
"""


import datetime
import itertools
import os
import re
import subprocess
import sys
import time


CONN_SETTINGS = {
    'server': 'pdxvmdevops01',
    'database': 'ReinsuranceSettlements',
    'quiet': True
}


def main():
    script = sys.argv[0]
    os.chdir(os.path.split(script)[0]) # initial working directory is the one where files were dragged from
    paths = sys.argv[1:]

    print(f"Testing connection to {CONN_SETTINGS['server']}.{CONN_SETTINGS['database']}...")
    if test_connection(**CONN_SETTINGS):
        raise RuntimeError(f"Error connecting to {CONN_SETTINGS['server']}.{CONN_SETTINGS['database']}. Check names and try again.")

    version = make_temp_tables(**CONN_SETTINGS)
    time.sleep(1)
    print(f'Created temp tables with version number {version}...')

    file_settings = [get_load_settings('format_files', p, version) for p in paths]    
    for fs in file_settings:
        if os.path.isdir(fs['path']):
            print(f'Must load individual files, not directories ({fs["path"]} is a directory).')
            return
        if bcp(**fs, **CONN_SETTINGS, errorfile='err.txt'):
            print(f'Error loading {fs["path"]}, err.txt file may contain additional detail.')
            return
        else:
            print(f'Loaded {fs["path"]}...')

    results = validate(**CONN_SETTINGS, table_version=version) # this will print validation messages


def test_connection(server, database, quiet, **kwargs):
    query = 'select 1 as val into #validator_connection_test'
    cmd = ['sqlcmd', '-S', server, '-d', database, '-Q', query]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    output = str(proc.stdout, encoding='utf8').split('\r\n')
    errors = any('error' in line or 'Error' in line for line in output)
    if errors or not quiet:
        for line in output:
            print(line)
    return proc.returncode or errors


def get_load_settings(format_file_dir, relpath, version):
    """
        Expects absolute paths
    """
    end_of_name = relpath.split("_")[-1]
    file_type = os.path.splitext(end_of_name)[0]
    path = os.path.abspath(relpath)
    fmt_path = os.path.join(format_file_dir, file_type + '.fmt')
    settings = {'table': file_type + '_' + version,
                'path': path, 
                'format_file': fmt_path}
    return settings


def make_temp_tables(server, database, quiet, **kwargs):
    """ Create global temp tables so that sqlcmd has a place to put data
        that won't mess up anything else that is running.
    """
    time_ms = datetime.datetime.now().timestamp()
    version = str(time_ms).split('.')[0]
    tmp_table_sql = f"""
        SELECT * INTO ##alfa0_{version} FROM Stage.Alfa0 WHERE 1=2;
        SELECT * INTO ##avrf2_{version} FROM Stage.AVRF2 WHERE 1=2;
        SELECT * INTO ##avrf3_{version} FROM Stage.AVRF3 WHERE 1=2;
        SELECT * INTO ##modeling4_{version} FROM Stage.Modeling4 WHERE 1=2;
        SELECT * INTO ##policies4_{version} FROM Stage.Policies4 WHERE 1=2;
        SELECT * INTO ##policies5_{version} FROM Stage.Policies5 WHERE 1=2;
        SELECT * INTO ##transactions3_{version} FROM Stage.Transactions3 WHERE 1=2;
        WAITFOR DELAY '00:10:00';
    """
    # Note that we're not capturing the process created with Popen because we're planning for it to when the script is closed (and sqlcmd is therefore terminated).
    subprocess.Popen(['sqlcmd', '-S', server, '-d', database, '-Q', tmp_table_sql], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return version


def bcp(table, path, format_file, server, database, errorfile, quiet, **kwargs):
    errorpath = os.path.abspath(errorfile)
    cmd = ['bcp', '##' + table, 'in', path, '-f', format_file, '-S', server, '-d', database, '-T', '-m', '0', '-e', errorpath] # this is the one place where max_errors=1 implies quit on the first....
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    output = str(proc.stdout, encoding='utf8').split('\r\n')
    errors = any('Error =' in line for line in output)
    if errors or not quiet:
        for line in output:
            print(line)
    return proc.returncode or errors


def validate(server, database, quiet, table_version, **kwargs):
    sql = f"""
        SET XACT_ABORT ON
        SET NOCOUNT ON

        BEGIN TRANSACTION
        BEGIN TRY

            DELETE FROM stage.ALFA0
            DELETE FROM stage.AVRF2
            DELETE FROM stage.AVRF3
            DELETE FROM stage.Modeling4
            DELETE FROM stage.Policies4
            DELETE FROM stage.Policies5
            DELETE FROM stage.Transactions3

            INSERT INTO stage.ALFA0 SELECT * FROM ##ALFA0_{table_version}
            INSERT INTO stage.AVRF2 SELECT * FROM ##AVRF2_{table_version}
            INSERT INTO stage.AVRF3 SELECT * FROM ##AVRF3_{table_version}
            INSERT INTO stage.Modeling4 SELECT * FROM ##Modeling4_{table_version}
            INSERT INTO stage.Policies4 SELECT * FROM ##Policies4_{table_version}
            INSERT INTO stage.Policies5 SELECT * FROM ##Policies5_{table_version}
            INSERT INTO stage.Transactions3 SELECT * FROM ##Transactions3_{table_version}

            EXEC stage.TranslateToCurrentVersion
            IF XACT_STATE() <> 1
                RAISERROR('dbo.DataLoadLive stage.TranslateToCurrentVersion error', 16, 1)

            EXEC stage.Validate 1 -- include output
            IF XACT_STATE() <> 1
                RAISERROR('dbo.DataLoadLive stage.Validate error', 16, 1)
        END TRY
        BEGIN CATCH
            PRINT ERROR_MESSAGE()
        END CATCH
        IF @@TRANCOUNT > 0 
            ROLLBACK TRANSACTION
    """
    # sqlcmd args are -S for server, -d for database, -j to show PRINT messages, and -Q for run query then exit
    proc = subprocess.run(['sqlcmd', '-S', server, '-d', database, '-r1', '-Q', sql], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    output = str(proc.stdout, encoding='utf8').split('\r\n')
    errors = len(output) > 11
    if errors or not quiet:
        print('Validation messages (first 10 errors)...\n')
        for line in take(21, output):
            line = re.sub('^.*\[SQL Server\]', '', line)
            print(line)
    elif not errors:
       print('No validation issues found.')
    return



def take(n, iterable):
    "Return first n items of the iterable as a list"
    return list(itertools.islice(iterable, n))
	
	
	
if __name__ == '__main__':
    try:
        main()
        print('\nClose the window or press CTRL+C to exit.')
    except Exception as e:
        print(e)
        input('\nERROR!\nClose the window to exit.')
        sys.exit()
