# Python Pitch Shift Script
### Introduction
This script has been developed in the context of the TU Delft Bachelor's Honours Programme for research purposes. The goal of the research was to investigate the effects on adult A.S.R. of artificially created child speech data from adult speech using pitch shift.

The pitch shift script was created using [PySox](https://pypi.org/project/pysox/) and [CREPE](https://pypi.org/project/crepe/).

### Internal workings
The script works by taking in adult speech, choosing a random target frequency in the interval [250Hz, 300Hz] and then using CREPE to determine how far we have to shift to get the file in the target frequency. The shift is done by PySox and verified by using CREPE again. Both the original and shifted properties are saved in an Excel sheet at the end of the script.

### Usage
To use this script, you can run it by using the following command:
```
python main.py source_dir target_dir excel_sheet_name
```

### Issues
Because not all adult speech is fit to be a source for pitch shifting, it is possible that this script creates "extreme" cases where the quality is low. These files can be separated as they have the "pse" tag instead of "ps", or simply by investigating the "extreme" column in the Excel sheet. In general, woman speech tends to be a better source as they are closer to child speech than males.

## Credits
Dr. O.E. Scharenborg  
Dr. T. Patel  
S.M.Z. Van den Broeck
