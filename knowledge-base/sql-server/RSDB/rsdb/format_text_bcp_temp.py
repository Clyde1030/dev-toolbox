import argparse
import re
import os
import csv
from datetime import datetime
from dateutil import parser
import shutil
import logging

LOGGER = logging.getLogger()

def main():
    args = get_args()
    configure_logger(LOGGER, args.loglevel)
    LOGGER.info('Formatting initializing...')
    data_dir = args.data and re.sub(r'[\'"]$', '', args.data) # remove trailing single or double quotes (happens when directories are passed in as "J:\whatever\" including the trailing slash)
    save_dir = args.save and re.sub(r'[\'"]$', '', args.save)
    LOGGER.debug('parsed args')
    if os.path.exists(save_dir):
        shutil.rmtree(save_dir)
    os.mkdir(save_dir)
    LOGGER.info('Cleared previous directory')
    h = getFileList(data_dir, save_dir, args.pattern)
    LOGGER.info('List of files to format and save locations generated')
    formatFiles(h)
    LOGGER.info('Formatting complete.')

def get_args():
    parser = argparse.ArgumentParser(description="RSDB control program")
    parser.add_argument('--data', default='', help="Directory that contains files to be formatted. \
        Default location: J:\\Acctng\\QuarterClose\\2019\\Q1\\Assumed Settlements\\Data")
    parser.add_argument('--save', default='', help="Directory that contains save locations for formatted files. \
        Default location: J:\\Acctng\\QuarterClose\\2019\\Q1\\Assumed Settlements\\bcpdata")  
    parser.add_argument('--pattern', default='(?!.*superseded)^.*csv$') 
    parser.add_argument('--loglevel', default='INFO', help="Minimum level of log output (DEBUG/INFO/WARNING/ERROR/CRITICAL), default is 'INFO'")
    args = parser.parse_args()
    return args

def getFileList(data_dir, save_dir, regex):
    LOGGER.debug('Pulling data files to be formatted...')
    f = []
    s = []
    for dirpath, _, filenames in os.walk(data_dir):
        data_paths = [os.path.join(dirpath, f) for f in filenames]
        fileLoc = [os.path.join(data_dir, p) for p in data_paths if re.search(regex, p)] # can't use abspath since relative to cwd
        f.extend(fileLoc)
    
    LOGGER.debug('Generating locations to save formatted data files...')
    for p in f:
        t = '0'
        save=[save_dir + p.replace(data_dir,'')]
        save2 = save[0].lower()
        if save2.endswith('_transactions.csv'):
            t = '3'
        elif save2.endswith('_policies.csv'):
            t = '4'
        elif save2.endswith('_modeling.csv'):
            t = '4'
        elif save2.endswith('_avrf.csv'):
            t = '2'
        save = os.path.splitext(save[0])[0] + t + '.csv'
        s.append(save)
    r = list(zip(f,s))
    return r

def formatFiles(fileList):
    LOGGER.debug('Beginning formatting...')
    for src, dst in fileList:
        dst_dir = os.path.dirname(dst)
        os.makedirs(dst_dir, exist_ok=True)
        src = src.lower()
        try:
            if src.endswith('_transactions.csv'):
                LOGGER.debug('Formatting ' + src + ' as transactions file')
                format(src, dst, fmt_trans)
            elif src.endswith('_policies.csv'):
                LOGGER.debug('Formatting ' + src + ' as policies file')
                format(src, dst, fmt_pol)
            elif src.endswith('_modeling.csv'):
                LOGGER.debug('Formatting ' + src + ' as modeling file')
                format(src, dst, fmt_mod)
            elif src.endswith('_avrf.csv'):
                LOGGER.debug('Formatting ' + src + ' as avrf file')
                format(src, dst, fmt_avrf)
            else:
                LOGGER.warning(src + ' not formatted')
        except Exception as e:
            LOGGER.error(f'error processing file {src}', e)

        LOGGER.debug('Saving formatted file to ' + dst)

def format(path, dst, fmt_func):
    line_count=0
    with open(path, 'r', newline='') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        with open(dst, 'w', newline='') as new_csv_file:
            csv_writer = csv.writer(new_csv_file,delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            for row in csv_reader:
                new_row = fmt_func(row)
                csv_writer.writerow(new_row)
                line_count +=1
    LOGGER.debug(f'processed {line_count} lines.')

def fmt_pol(row):
    row[0] = formatDate(row[0])
    row[4] = formatDate(row[4])
    if row[7] != '':
        row[7] = formatMoney(row[7])
    if row[8] != '':
        row[8] = formatMoney(row[8])
    if row[9] != '':
        row[9] = formatMoney(row[9])
    if row[10] != '':
        row[10] = formatMoney(row[10])
    if row[11] != '':
        row[11] = formatMoney(row[11])
    if row[12] != '':
        row[12] = formatMoney(row[12])
    if row[13] != '':
        row[13] = formatMoney(row[13])
    if row[14] != '':
        row[14] = formatMoney(row[14])
    row[19] = formatQS(row[19])
    row[20] = formatMoney(row[20])
    row[21] = formatMoney(row[21])
    row[22] = formatMoney(row[22])
    row[23] = formatMoney(row[23])
    row[24] = formatMoney(row[24])
    row[25] = formatMoney(row[25])
    row[26] = formatMoney(row[26])
    row[27] = formatMoney(row[27])
    return row

def fmt_trans(row):
    row[0] = formatDate(row[0])
    row[8] = formatMoney(row[8])
    row[9] = formatDate(row[9])
    return row

def fmt_mod(row):
    row[0] = formatDate(row[0])
    row[6] = formatRating(row[6])
    if row[10] != '':
        row[10] = formatRating(row[10])
    row[11] = formatMoney(row[11])
    row[12] = formatMoney(row[12])
    if row[14] != '':
        row[14] = formatDate(row[14])
    row[15] = formatMoney(row[15])
    return row

def fmt_avrf(row):
    row[0] = formatDate(row[0])
    row[8] = formatMoney(row[8])
    row[9] = formatDate(row[9])
    return row

def formatDate(date):
    asDate = parser.parse(date)
    asString = asDate.strftime('%Y-%m-%d %X')
    return asString

def formatMoney(number):
    asFloat = round(float(number),4)
    asString = '{:.4f}'.format(asFloat)
    return asString

def formatQS(number):
    asFloat = round(float(number),8)
    asString = '{:.8f}'.format(asFloat)
    return asString

def formatRating(number):
    asFloat = round(float(number),7)
    asString = '{:.7f}'.format(asFloat)
    return asString

def configure_logger(logger, level):
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    log_level = level.upper()
    logger.setLevel(log_level)

if __name__ == '__main__':
    main()
