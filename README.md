PyOfficeGrep
===

![Software Version](http://img.shields.io/badge/Version-v0.0.2-green.svg?style=flat)
![Python Version](http://img.shields.io/badge/Python-3.x-blue.svg?style=flat)
[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

[Japanese Page](./README.ja.md)

## Overview
Run a grep search on the Office files.

## Version
v0.0.2

## Requirements
### When using `office_grep.py`
- Python 3.x
- PyWin32
- colorama

It's all included in [Anaconda3](https://www.anaconda.com/).

### When using `office_grep.exe`
Nothing.

Note:
- You need `PyInstaller` to generate an exe file using `to_exe.bat`.

## License
[MIT License](./LICENSE)

## Spec and Limitations
- Windows only.
- Currently, only Excel and Word files (xls, xlsx, xlsm, doc, docx, docm) are supported.
    - I would like to support PowerPoint files (ppt, pptx, pptm) in the future, but this has been put off because we haven't been forced to.
- To speed up the search for cells in the Excel file, a slight modification was made.

## How to use
Run `office_grep.py` or `office_grep.exe` with arguments.

```
$ python office_grep.py -h
usage: office_grep.py [-h] [-v] [--type TYPE] [--word WORD]
                      [--recursive RECURSIVE] [--ignorecase IGNORECASE]
                      [--regex REGEX] [--parallel PARALLEL]
                      query dirpath

Run a grep search on the Office files.

positional arguments:
  query                 Search query
  dirpath               Target directory path

optional arguments:
  -h, --help            show this help message and exit
  -v, --version         show program's version number and exit
  --type TYPE           Target file types (The following letter combinations:
                        "E":Excel, "W":Word, "P":PowerPoint)
  --word WORD           Match whole words only
  --recursive RECURSIVE
                        Search the directory recursively
  --ignorecase IGNORECASE
                        Case insensitive
  --regex REGEX         Regex
  --parallel PARALLEL   Maximum number of parallel processing
```

### Required parameters

| Parameter | Explanation           | Value Format |
|-----------|-----------------------|--------------|
| query     | Search query          | String       |
| dirpath   | Target directory path | Path string  |

### Optional Parameters

| Parameter  | Explanation                           | Value Format                                                         | Initial Value |
|------------|---------------------------------------|----------------------------------------------------------------------|---------------|
| type       | Target file types                     | The following letter combinations:<br/>E:Excel, W:Word, P:PowerPoint | EWP           |
| word       | Match whole words only                | Boolean (True or False)                                              | False         |
| recursive  | Search the directory recursively      | Boolean (True or False)                                              | True          |
| ignorecase | Case insensitive                      | Boolean (True or False)                                              | False         |
| regex      | Regex                                 | Boolean (True or False)                                              | True          |
| parallel   | Maximum number of parallel processing | Integer                                                              | 1             |

Note:  
- parallel
    - Specify 1 to disable parallel processing.
    - It becomes impossible to stop by interrupts such as Ctrl+C when parallelized.
    - Currently, 1 is recommended because the speed improvement cannot be expected even if it is parallelized.

#### How the values are adopted
The values are adopted in the following order of priority:
1. Command line arguments.
2. Configuration file. (`setting.ini` in the current directory.)
3. Initial value.


## Content of results

![result](https://user-images.githubusercontent.com/64964079/86204256-0028ee80-bba2-11ea-8093-1cb48b20acb9.png)

### Line in the file path
A file path is displayed with the current file number and the total number of all files that have been processed.  
The first character indicates the file type. (E:Excel, W:Word, P:PowerPoint)

### Line in the result
This is displayed if there are search results for that file.

The first half is information block, about the location of the search result.  
The second half is the text block, including the before and after of the search results. The tokens found in the search will be highlighted.

The following is a description of the information blocks for each file type.

#### Excel
Kind (front of a parenthesis):

| Kind    |
|---------|
| Cell    |
| Shape   |
| Comment |

Details (in parentheses):

| Details | Explanation   |
|---------|---------------|
| Sheet   | Sheet name.   |
| Address | Cell address. |
| Name    | Object name.  |

#### Word
Kind (front of a parenthesis):

| Kind    |
|---------|
| Text    |
| Table   |
| Shape   |
| Comment |

Details (in parentheses):

| Details | Explanation                                     |
|---------|-------------------------------------------------|
| Page    | Page number.                                    |
| Line    | Line number in page.                            |
| Cell    | Position of a cell in the table. (row, column). |
| Name    | Object name.                                    |

#### PowerPoint
WIP
