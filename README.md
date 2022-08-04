# Read Email from Gmail and update mail on Google Sheet Using Python
---
This python script is used to read an unread email with a specific label from Gmail inbox and get desired information from the email and put it on a google sheet.The project uses google Sheets and Gmail API. If this script runs periodically on the cloud like Amazon EC2 it will read mail and put information from email in google Sheets

### Applications
The application of this script is endless. Following are some real-world uses of this script.

- You want to create your new job listing website. You can go to many job searching websites and subscribe using your email address.Websites send new job listings on your email. This script fetches this email and forms a list of the new job listings and puts it in the spreadsheet.Now from a spreadsheet, you can attach your new website. Your website updates new job listings from the spreadsheet.

- You are a crypto trader and set an email alerts on the price movement of specific crypto assets. This script is running on amazon EC2 24 hours / 7 days and fetches alerts on email and adds a buy signal on the spreadsheet. From the spreadsheet, you can connect your bot.

- You are a content writer and want to get content ideas from famous authors/websites. You subscribe to authors/websites to get an email notification on their new article.This script makes a list of article and their links which make it easier for you to get updated with content.

- You are a marketing researcher and subscribe to email notifications to marketing research websites. This script list all notification in the google sheet.

### Implementation
The script is written by supposing that Gmail and google Sheets are of the same google account.
To implement the script we need the following prerequisites:
- Python 3
- The [pip](https://pypi.python.org/pypi/pip) package management tool
- A Google Cloud Platform project with the API enabled. To create a project and enable an API, refer to [Create a project and enable the APIs](https://developers.google.com/workspace/guides/create-project)(both google sheet and gmail api)

- Authorization credentials for a desktop application.Create OAuth 2.0 client. To learn how to create credentials for a desktop application, refer to [Create credentials](https://developers.google.com/workspace/guides/create-credentials)
- Now [Create testing User](https://console.cloud.google.com/apis/credentials/consent) for OAuth  by clicking test user.Test user is create by simply providing email address of google account(Test user is also of same google account of which Gmail and sheet is use in the project)

For step 3 and 4 you also refer this [article](https://codehandbook.org/how-to-read-email-from-gmail-api-using-python/). After step 2 and 3 you get credentials.json file which you paste in this project folder.

Now install the Google client library for Python, run the following command:

```console
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```
Before running, read_mail.py script put spreadsheet id in script. Also change "search_string" to filter particular email from unread emails

## Developer
If you are not familiar with python and want to modify this script as per your particular application feel free to contact

Contact :msalmanashraf22@gmail.com