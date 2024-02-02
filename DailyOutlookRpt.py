import win32com.client
from pathlib import Path
import re

def get_reports_folder():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Get the Inbox folder
    inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder
    
    # Navigate to the "Reports" folder under the Inbox
    reports_folder = None
    for folder in inbox.Folders:
        if folder.Name == "Reports":
            reports_folder = folder
            break
    
    # Release the Outlook object
    del outlook
    
    return reports_folder

def process_reports_folder(folder):
    if folder is None:
        print("Reports folder not found.")
        return
    
    output_dir = Path.cwd() / "Output"
    output_dir.mkdir(parents=True, exist_ok=True)

    messages = folder.Items
    for message in messages:
        subject = message.subject
        subject_clean = re.sub('\W+'," ", subject )
        body = message.body
        attachments = message.attachments
        received_time = message.receivedtime
        formatted_received_time = received_time.strftime("%Y_%m_%d")
        final_format = formatted_received_time.replace(":",".")
        
        #Create the folder from email subject & date 
        folder_name = f"{subject_clean}_{final_format}"


        target_folder = output_dir / folder_name
        target_folder.mkdir(parents=True, exist_ok=True)

        #This is if we want to pull the body of the email into file explorer 
        #Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))


        for attachment in attachments:
            attachment.SaveAsFile(target_folder / str(attachment))



reports_folder = get_reports_folder()
process_reports_folder(reports_folder)


 # Path to the Python file you want to run
#file_to_run = "C:\Users\xavie\Desktop\Outlook Py\combine_excel.py"

# Use exec to run the other Python file in the same process
#exec(open(file_to_run).read()) 
