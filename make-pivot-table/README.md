# make pivot table script

## Prerequisites

1. Install Anaconda
* [Anaconda For Windows / Python 2.7 64-bit version] (https://www.continuum.io/downloads#windows)

## How to run command
* Run Anaconda Prompt

```dos
(C:\Anaconda2) C:\Users\john>dir /O-D /w
 C 드라이브의 볼륨에는 이름이 없습니다.
 볼륨 일련 번호: D8FC-8195

 C:\Users\john 디렉터리

[..]
[.]
mk_pv.py
output.xlsx
...

(C:\Anaconda2) C:\Users\john>python mk_pv.py --help
usage: mk_pv.py [-h] [-v] [--visible] [-s SUFFIX] [--pt_ro PT_RO]
                [--pt_co PT_CO] [--pc_ro PC_RO] [--pc_co PC_CO]
                [--pc_type PC_TYPE] [--pc_width PC_WIDTH]
                [--pc_height PC_HEIGHT] [--pc_y_max PC_Y_MAX]
                input_file_path sheet_indices [sheet_indices ...]

making pivot tables & charts

positional arguments:
  input_file_path       input file path
  sheet_indices         target sheet indices

optional arguments:
  -h, --help            show this help message and exit
  -v, --verbosity       increase output verbosity
  --visible             setting Excel visible
  -s SUFFIX, --suffix SUFFIX
                        setting save file suffix
  --pt_ro PT_RO         setting pivot table row offset
  --pt_co PT_CO         setting pivot table column offset
  --pc_ro PC_RO         setting pivot chart row offset
  --pc_co PC_CO         setting pivot chart column offset
  --pc_type PC_TYPE     setting pivot chart type
  --pc_width PC_WIDTH   setting pivot chart width
  --pc_height PC_HEIGHT
                        setting pivot chart width
  --pc_y_max PC_Y_MAX   setting pivot chart y axis max

(C:\Anaconda2) C:\Users\john> python mk_pv.py output.xlsx 2
or
(C:\Anaconda2) C:\Users\john> python mk_pv.py --pc_type=xlAreaStacked output.xlsx 2 3

(C:\Anaconda2) C:\Users\john>dir /O-D /w
 C 드라이브의 볼륨에는 이름이 없습니다.
 볼륨 일련 번호: D8FC-8195

 C:\Users\john 디렉터리

[..]
[.]
output_pv.xlsx
mk_pv.py
output.xlsx
...

```

