# Python x MySQL x Excel by Zinglecode
 
Example Python codes that do the processes between MySQL database and Excel spreadsheet files.

## Setup database table

![Database table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel-1/dev/snapshots/db-table-structure.png?token=AC32DU3FQXPZQTWUTPCOUTTAVDU4I "Database table structure")
*Structure*

![Database table sample data](https://raw.githubusercontent.com/potchangelo/python-mysql-excel-1/dev/snapshots/db-table-data.png?token=AC32DU2GRDIWEPS42KOMO6DAVDU6G "Database table sample data")
*Sample data*

## Install Python 3 and pipenv

1. Download Python 3 installation file from https://www.python.org/

2. Install pipenv as global package by this command.

```
pip install pipenv
```

Note: for macOS with pre-installed Python 2, use pip3 instead of pip.

## Install and run project by PyCharm

1. Download this project.

2. Open PyCharm and choose File -> Open... -> Then select project folder

3. Fix any warning recommended by PyCharm.

4. Make sure that every project packages is installed, by open PyCharm Terminal and type command.

```
pipenv install
```

5. Open "main.py", right click on code area, select Run 'main'

6. Program will inform you, to type one file name from "scripts" folder to be run.

## Install and run project by CLI

1. Download this project

2. Open Terminal or Command Prompt at project folder, then install packages.

```
pipenv install
```

3. Activate pipenv environment.

```
pipenv shell
```

4. Run "main.py"

```
python main.py
```

5. Program will inform you, to type one file name from "scripts" folder to be run.
