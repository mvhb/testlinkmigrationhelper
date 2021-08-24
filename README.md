# TestLink migration to Excel (xlsx) - helper

## Main goal
```
Generate a EXCEL (xlsx) file given a exported XML from TestLink.
The exported XML must be folder by folder, it's not possible to convert all testcases at once.  
```

## Installation
```
pip install -r requirements.txt
```

## Pre condition
```
The XML file must be present on TestLinkMigrationHelper/xml folder.
```

## Run Script
```
python excel_migrator
xlm_name_example.xml

To quit, instead of xml filename, digit "quit".
```
