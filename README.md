# Python x MySQL x Excel by Zinglecode
 
Example Python codes that do the processes between MySQL database and Excel spreadsheet files.

## YouTube videos

- [MySQL #01](https://www.youtube.com/watch?v=axraNvtHjO4)
- MySQL #02 is coming soon

## Setup database table

![products table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-products-structure.jpg "products table structure")
*products table structure*

![product_notes table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-product-notes-structure.jpg "product_notes table structure")
*product_notes table structure*

![categories table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-categories-structure.jpg "categories table structure")
*categories table structure*

![hashtags table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-hashtags-structure.jpg "hashtags table structure")
*hashtags table structure*

![products_hashtags table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-products-hashtags-structure.jpg "products_hashtags table structure")
*products_hashtags table structure*

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
