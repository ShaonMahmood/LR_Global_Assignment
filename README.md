# LR_Global_Assignment
The project is tested in a ubuntu environment

## Compulsory Steps
- Open command prompt
- Make sure python version >=3.6 is installed and git installed
- Make a project folder and cd into it. for example: `mkdir Myproject && cd Myproject`
- git clone the project from here and cd into the cloned folder.
- Create a python3 virtual environment using command `python3 -m venv env`
- Enter into the virtual environment using `source env/bin/activate`
- install the requirements using `pip install -r requirements.txt`

## Steps to run the automation script
To run the automation script you have two options,either with verbosity or without verbosity. two required positional arguments is required.viz: inputfile and outfile. the format is `python automation.py <input_file> <output_file> --verbosity`. for example:

- Since out input and output file is in data folder, we can run the script using command with verbosity `python automation.py "data/Compiled Index.xlsx" "data/final_output.csv" --verbose`
- if you don't want verbosity then command will be `python automation.py "data/Compiled Index.xlsx" "data/final_output.csv"`
- lastly after the running the script you will find your output in data/final_output.csv file or in the file you specified in command line argument.

## Steps to run the notebook
the notebook containes the detailed explanation with output. the steps to run it are: 

- run jupyter notebook by simply writing `jupyter-notebook` into the prompt.A browser will pop up.
- open read into dataframe.ipynb from notebook. and you are off to go.
- for testing with changing excel sheets, change the  xlsx file on data folder. (converted to xlsx file from xlsm for simplicity)
- the final output result will be found on data/final_output.csv
- the solution has many lackings as the time was short but its works. So Enjoy!!
