# Test profile tool

This tool can generate Excel files baesd on profile txt data from the  halo2 project.

## How to run
* prepare libs
```bash
pip3 install openpyxl
```
* Run
  For now, you have to modify the target input and output in the file. (Later to support pass them with params.)

Suppose input_file(./xxx.txt), output_file(../xxx.xlsx)
```
python3 profiling_analysis.py input_file output_file
```
