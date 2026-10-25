@SETLOCAL ENABLEDELAYEDEXPANSION & python -x "%~f0" %* & (IF ERRORLEVEL 1 PAUSE) & EXIT /B !ERRORLEVEL!
"""
    This is a quick-and-dirty script to run the validation proc on files that are dragged and dropped.
    Hence, it is designed to work without installing any dependencies. Hence, it has some workarounds.
    and minimal error handling!
"""


import datetime
import os
import subprocess
import sys
import time


CONN_SETTINGS = {
    'server': 'pdxvmdb11',
    'database': 'ProductionAndOverride',
    'quiet': True
}


def main():

    script = sys.argv[0] 
    os.chdir(os.path.split(script)[0])
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

def make_temp_tables(server, database, quiet, **kwargs):
    """ Create global temp tables so that sqlcmd has a place to put data
        that won't mess up anything else that is running.
    """
    time_ms = datetime.datetime.now().timestamp()
    version = str(time_ms).split('.')[0]
    tmp_table_sql = f"""
        SELECT * INTO ##_{version} FROM dbo.OverrideStaging WHERE 1=2;
        WAITFOR DELAY '00:10:00';
    """
    # Note that we're not capturing the process created with Popen because we're planning for it to when the script is closed (and sqlcmd is therefore terminated).
    subprocess.Popen(['sqlcmd', '-S', server, '-d', database, '-Q', tmp_table_sql], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return version

def get_load_settings(format_file_dir, relpath, version):
    """
        Expects absolute paths 
    """
	# relpath: C:\Users\yu-shenglee\Desktop\Monarch immigration\Validator\O-1111-LIF-2022-12-123V.csv
    path = os.path.abspath(relpath)
    fmt_path = os.path.join(format_file_dir, 'OverrideStaging.fmt')
    settings = {'table': 'dbo.OverrideStaging',
                'path': path, 
                'format_file': fmt_path}
    return settings

def bcp(table, path, format_file, server, database, errorfile, quiet, **kwargs):
    errorpath = os.path.abspath(errorfile)
    cmd = ['bcp', table, 'in', path, '-f', format_file, '-S', server, '-d', database, '-T', '-m', '0', '-e', errorpath] # this is the one place where max_errors=1 implies quit on the first....
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    output = str(proc.stdout, encoding='utf8').split('\r\n')
    errors = any('Error =' in line for line in output)
    if errors or not quiet:
        for line in output:
            print(line)
    return proc.returncode or errors

if __name__ == '__main__':
    try:
        main()
        print('\nClose the window or press CTRL+C to exit.')
    except Exception as e:
        print(e)
        input('\nERROR!\nClose the window to exit.')
        sys.exit()
