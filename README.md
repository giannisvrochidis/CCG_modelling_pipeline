# CCG energy modelling pipeline

## Getting Started

### Prerequisites
 - Windows 10 or later
 - Microsoft Office 2016 or later
 - Python 3 [(ex. 3.9.12)](https://www.python.org/downloads/release/python-3912/)

### Setup your environment 

1. Make sure you have installed all the prerequisites.

2. Run the `setup.bat` script.

3. You need to change your Windows `regional format` from the `regional settings` to `English (United Kingdom)`.

4. You need to change your Office language to `English` if it is not already.

5. You need to `enable macros` in Microsoft Excel.

### Configure the pipeline

Open the `settings.json` configuration file and enter your configuration settings.

### Run the application

1. Open a cmd and head to the project's directory.

2. Enter the virtual environment by running:
```
    pipenv shell
```

3. To run the whole pipeline run:
```
    python main.py
```
or alternatively run a specific script.

```
    python maed.py
    python osemosys.py
    python flextool.py
```
