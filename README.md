# dclmic-export

Python library to facilitate exporing pandas dataframes to excel sheets.

## Installation

``` bash
pip install git+https://github.com/Dallas-College-LMIC/dclmic-export
```

## Usage


``` python
import dclmic-export

dclmic-export.save_dfs_as_xl(
    list_of_frames=[df1, df2, df3]
    path="./output/"
    file_name="my_dfs"
    sheet_names=["Dataframe 1", "Dataframe 2", "Dataframe 3"]
)
```

See the docstring for save_dfs_as_xl for details on all of the parameters.


