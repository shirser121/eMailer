# eMailer - send emails from Excel file!

### This is a simple emailer that can be used to send emails to a list of people.

#### How this is working?
The package get an email, password and path to Excel file.
The package will read the Excel file and will email each person in the list.  
You can also add a subject and a message to the email.
You can use another parameters from the Excel file, just add {parameter} to the message.

#### examples:
```python
from emailer import emailer
emailer(email=xxxxxxx@gmail.com, password=xxxxxxx, spreadsheet=path_to_excel_file)
```

You can also use this package as CLI
```CommandLine
python emailer.py -e xxxxxxx@gmail.com -p xxxxxxx -s path_to_excel_file
```

You can also use config.txt to configure the package.
This is the options of the config file:  
- EMAIL=
- EMAIL_PASSWORD=
- FROM=
- SUBJECT=
- EXCEL_FILE_NAME=
- MESSAGE_FILE_NAME=


