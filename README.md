# Purpose of the project

This project was made as an assessment for a software engineering position, and its purpose is to parse a excel file in a undesired format and then write it to a more desirable and easier to read format in a new file.

# Expected input and output

The expected input is a path to the excel file to be parsed.
The expected output is a new excel file.

# Installation steps 
To install the required libraries, you need to have python and pip installed on your machine, and then run 
```
pip install -r requirements.txt
```

# Instructions to run the project
To run the project, simply run the `main.py` passing the path to the input excel file
```
python main.py input.xlsx
```

# Instructions to run tests
To run the tests execute the following command
```
pytest tests.py
```
And if you want to run with coverage
```
pytest --cov=main tests.py
```

# Use cases and edge cases covered in the code
The program covers the default use case, where the input is in the expected format, being able to parse any number of sites and any number of date data points. It also handles dates which has only a subset of metrics.

# Known limitations
Because this programs parses the whole excel file at once, it is limited by the available memory in the system, and will not work on excel files that are overly large.