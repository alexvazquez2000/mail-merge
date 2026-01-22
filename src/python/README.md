#To Install dependencies
    #create venv first
    pip install -r requirements.txt

#To freeze the requirements run the command:

    pip freeze > requirements.txt

# create venv environment

    python -m venv venv
    #on windows use:
    .\venv\Scripts\activate.bat
    #or on linux/Mac use:
    . venv/Scripts/activate
    pip list  #that showed an almost empty list
    #install the requirements
    pip install -r requirements.txt
    pip freeze > requirements.txt
    #now list shows about 20 packages
    pip list


# Run locally with python

    python app.py
