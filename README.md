# General Information
lexmark_ink_reviewer.py is a script that iterates through lexmark printers web interfaces and scrap his data for reporting purposes, this script was tested with the following printers:
- MX431adn
- CX522ade
- MX611dhe
- MX622adhe
- MX622adn
# Requirements
- Python 3.12.x or higher
- Other requirements specified on requirements.txt
# Setup
Ensure that python.exe is added on system environment variables for executing the script on task scheduler.

Create a directory to install the following inside that directory:
- *Python virtual environment*: execute the command `python -m venv your_venv_name` for creating a folder with all the dependencies needed.
  Activate your venv on venv/Scripts/activate, once that done, install the dependencies specified on `requirements.txt` with `pip -r install requirements.txt`
- *Clone git repo*: Clone `https://github.com/xChebix/Lexmark_ink_reviewer.git`

On script lexmark_ink_reviewer.py, on variable filename, line 282 aprox. needs a modification, for the script to work, create a directory for storing the reports made by the tool, copy the path of the new directory and add it into with this format:
```
filename = fr"[folderpath]\printer_report {date_str}.pdf"
```

*Note:* The filename can be changed but there is needed the "{date_str}" string for assigning the date to the file.

- *.env*: you need to create a .env file for defining environment variables that are used on the script, these variables contain, user mail, password, recipients and copy mails, this is made this way for security purposes, the format for this file is this:
dotenv.env:
```
EMAIL_USER=examplemail@gmail.com
EMAIL_PASSWORD=emailpassword
HOST_MAIL=mail.organization.com
RECIPIENT=receiver@mail.com
EMAIL_CC_LIST='["mail1@example.com","jjauregui@agropartners.com.bo"]'`
```

*Note*: This is the format for the .env file which needs to be located on project folder that was cloned.

# Documentation
### Extract_Ips()
This functions uses pandas for extracting data from excel, it extracts an ip list from the third column and extract the area which belongs on the fourth column, then adds it into a dictionary and returns it as a list of dictionaries.
```
for index, row in data.iterrows():
        ip_dict = {
            "ip": str(row.iloc[2]),
            "area": str(row.iloc[3])
        }
        ip_addresses.append(ip_dict)
    return ip_addresses
```

## Extract_Printers_Data()
This function uses the ip list from Extract_Ips(), then passes the ip to the following functions: Allinks_UI_Printer(), Onlyblack_UI_Printer() and Old_UI_Printer(), this functions search through the html the values of the printer and return them on a dictionary.
Then Extract_Printers_Data() appends each dictionary into list, there are two lists, a list with failed connections due to error 404 on the IP proportioned, and a list with all the values that were found of the printers.

## Allinks_UI_Printer(), Onlyblack_UI_Printer() and Old_UI_Printer()
This functions receives an ip as a parameter and find through the html the following values, ink levels, standard bin level, imaging unit and maintenance kit level, these values are returned as a dictionary
This function uses selenium and beautiful soup as web scrappers

# Email_sender.py
## Send_Mail()
This function uses yagmail, dotenv,
On yagmail library you can specify values like sender receiver, which server to use for sending the mail via SMTP and port used, all this values are stored on environment variables for security purposes and for avoiding storing passwords or sensitive data on the code.
# Recommendation for improvement
- Add a module for printing a copy when making a report
- Add more interfaces for detecting the status of more printers
