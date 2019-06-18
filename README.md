## nonverbal queue converter

conversion script for 'nonverbal queue' excel datasets, which are grouped as input by **source** (citation) and colored by **queue category**.

The output of the script is a re-grouping of the queues by **queue category**, with the option to also include the **source**. 

## installation:

1. Either of

    A. download this repository
    
    B. download the source file, `group_by_color.py`

2. Make sure you have python installed, and that it is of a version greater than 3 (python 3.X)

3. Install the following dependencies: 
    - `styleframe`
    - `pandas`
    - `numpy`
    - `argparse`
    - `sys`

    All of these should be installed except `styleframe`, unless your python is really messed up.

    A nice easy way to install packaged is to do: `pip install <package_name>`, e.g. `pip install styleframe` to install styleframe. 

## usage:

easiest way to run it is as a command-line python program:
```
python group_by_color.py [-h] [-o OUTPUT] [-s] [-r ROWH] [-c COLW] input

positional arguments:
  input

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        defaults to 'grouped'
  -s, --include-sources
                        defaults to False
  -r ROWH, --row-height ROWH
                        defaults to 21
  -c COLW, --column-width COLW
                        defaults to 18 
```

For example, if you had a spreadsheet `mysheet.xlsx` that you wanted to convert into a spreadsheet `mysheet_converted.xlsx` with an excel row height of 30, you would run

``python group_by_color.py mysheet -o mysheet_converted -r 30``

(the `.xlsx` suffix is optional/assumed) <br>
To include the sources, add the `-s` flag:

``python group_by_color.py mysheet -o mysheet_converted -r 30 -s``

see the included `nonverbal_cues.xlsx` for a properly formatted cues file.