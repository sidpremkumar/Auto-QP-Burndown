## Auto-QP-Burndown :shipit:

### What is this? 
This is a function to automatically update a google doc based on the results of a JIRA filter 

### Setup 
You need to set up the following environmental variables: 
1. `JIRA_USER` :: Jira Username 
1. `JIRA_PW` :: Jira Password
1. `SPREADSHEET_ID` :: Spreadsheet ID (look at the URL of the spreadsheet)

### Running 
To run the program all you need to type is: 

    python main.py