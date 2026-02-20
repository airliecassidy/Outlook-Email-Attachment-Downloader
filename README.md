# Outlook Email Attachment Downloader


This tool was built to solve a recurring manual task: weekly reports arriving by email needed to be reliably saved to a shared drive without someone having to do it by hand each time. The script runs continuously in the background, checking the inbox on a configurable interval, downloading any new attachments from the expected sender, and sending an alert if the weekly email doesn't arrive when expected.


# Requirements 
- Windows only — requires a locally installed and running Outlook desktop client
- Python 3.x
- pywin32 library

# Logging
All activity is written to outlook_downloader.log in the working directory and printed to stdout. This includes successful downloads, skipped duplicates, missing email alerts, and any errors encountered.

# How It Works

1. Connects to the local Outlook instance via the Win32 COM API
2. Locates the target folder (searches shared mailboxes and inbox subfolders)
3. Filters emails by sender address and the expected delivery date
4.Skips any emails already recorded in processed_emails.json
5. Downloads all attachments to the configured destination, timestamping filenames if a conflict exists
6. Sends an SMTP alert if the expected weekly email is more than a day overdue
7. Sleeps for the configured interval, then repeats
