# Reinsurance Settlements Database README

This repository coordinates updates to the Reinsurance Settlements Database, so that we can reproduce the support for M Life's STAT and GAAP financial statements. Each quarter data is loaded from carrier settlement files, the Magnastar Database, TAI, and ALFA, and processed in one batch. Since settlements frequently are adjusted mid-quarter, the database is designed to be reset and re-run frequently throughout the quarter.


## How to set up on a new computer

The RSDB relies on third party Python packages which are managed through Pipenv. Install Pipenv from the network if you don't already have it (use the version on the network as the most recent versions have a different file format). Run `pipenv install` in the RSDB root folder, and Pipenv will install dependencies from where they are stored on the network.


## How to run the RSDB at quarter end

The program runs from the command line. Use the `run.bat` batch file in the root of the repository to run the program in the environment created by Pipenv. If you run it without any arguments it will give an overview of input options.

The most common command is `run.bat run %YYYY% %Q% --data %NETWORK_PATH%`, which will run using the ReinsuranceSettlements database on the local server, first reseting the database from the backup in the RSDB root folder and then running the refactor, load, and process instruction sets from the current quarter instruction file (which can be found in `etl\YYYY\QQ`). `run.bat fastforward` will run the same refactor, load, and process instruction sets without reseting the database at the start. The `--server` and `--database` options can be used to run the process on another server or against another database.


## How to create quarterly instruction files

The quarterly instruction files are in JSON format. Typically it is easiest to start a new file by copying the prior quarter's and making changes. New instruction sets can be added to the file and they will be followed if they are given as arguments. For example, `run.bat test_something 2020 3' will attempt to run the steps listed in the `test_something` section of `etl\2020\Q3\instructions.json`.


## How to create quarterly backups

Use `run.bat backup --backup %FILE_NAME%` to save a backup of the RSDB. After quarter close is complete, run a clean copy of the RSDB, create a backup named `ReinsuranceSettlements.bak`, put it in a zip file, and copy it to `J:\MLife\Backups\ReinsuranceSettlements\ReinsuranceSettlements\YYYYMMDD\`.
