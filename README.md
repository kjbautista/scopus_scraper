## Scopus Scraper

Fetch author-level metrics from Scopus using Pybliometrics. Inputs can be a text file (one name per line) or directly sent to the script using the --name argument (see Examples below). Output metrics are written to an Excel sheet.

### Setup

**Run the following in a bash (Linux) or a command prompt (Windows) terminal to setup the python environment:**
1) Setup virtual environment:
	```pip install virtualenv``` (if virtualenv is not already installed)
2) Create a new virtual environment (called venv here)
	```virtualenv venv```
3) Enter the virtual environment
	```source venv/bin/activate``` (Linux)
	OR
	```venv\Scripts\activate``` (Windows)
4) Install the required pacakges in the current environment.
	```pip install -r requirements.txt```

**Establish access to Scopus:**
1) Obtain an API key from the Elsevier Development Portal by following this [link](https://dev.elsevier.com/). You will need this on the first pass of the python script (main.py) to setup access to the Scopus API. 
2) Access may only work while on-campus on the UNC network.

### Usage

- **Text input** (one name per line):
	- Output defaults to `author_metrics.xlsx`.
- **Excel input** (`.xlsx`):
	- Output defaults to the same Excel file, written to a separate sheet (`metrics`).

### Examples

**Run the following in a bash (Linux) or a command prompt (Windows) terminal (***after activating the virtual environment***):**
- Use default files (input = names.txt, output = author_metrics.xlsx):
	- `python main.py`

- Specify input text file name and output file name:
	- `python main.py --input names.txt --output author_metrics.xlsx`

- Send names directly into the script:
	- `python main.py --name "Paul A Dayton" --name "Jane Doe" --output author_metrics.xlsx`


### Note:
- Example names in the names.txt file are current members of the [Dayton Lab](https://daytonlab.sites.unc.edu/?page_id=26) (as of Feb. 2026).
