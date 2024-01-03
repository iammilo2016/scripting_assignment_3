import subprocess #for running shell commands
import pandas as pd #for data manipulation and Excel file creation
import datetime #for timestamping the Excel file, and several
import win32com.client as win32 # modules for interacting with the Windows operating system.
import win32api #interact with the Windows API (Application Programming Interface) and retrieve information about the computer.
import win32file
import win32service #interact with the Windows Service Control Manager (SCM) and list all services.
import win32con # constants for the Windows API

def list_services(): # appends each service's short name, description, and status to a list
    try:
        services = []
        accessSCM = win32con.GENERIC_READ #access the SCM with read-only permissions
        hscm = win32service.OpenSCManager(None, None, accessSCM)
        typeFilter = win32service.SERVICE_WIN32 #filter for services that run in the Windows subsystem
        stateFilter = win32service.SERVICE_STATE_ALL #filter for services in any state

        statuses = win32service.EnumServicesStatus(hscm, typeFilter, stateFilter) #returns a list of tuples containing the short name, description, and status of each service

        for (short_name, desc, status) in statuses: #iterate through the list of tuples and append each service's short name, description, and status to a list
            services.append({
                'ServiceName': short_name,
                'Description': desc,
                'Status': status
            })

        return services
    except Exception as e:
        print('Error:', e)
        return []
    
def get_windows_updates():
    # Running PowerShell command to get update details
    process = subprocess.Popen(["powershell", "Get-HotFix"], stdout=subprocess.PIPE) #retrieves information about Windows updates. It processes the command's output, splits it into lines, and extracts the headers.
    result = process.communicate()[0].decode('utf-8') 
    lines = result.strip().split('\n')  # Split the data into lines
    headers = [h.strip() for h in lines[0].split()]  # Extract headers
    updates = []

    for line in lines[1:]:
        values = line.split(maxsplit=len(headers)-1)
        update_record = {headers[i]: values[i].strip() for i in range(len(headers))}
        updates.append(update_record)

    # Get additional system information using win32api
    computer_name = win32api.GetComputerName()
    for update in updates:
        update['ComputerName'] = computer_name

    return updates

def create_excel_file(update_info, filename, ):
    try:
    # Create DataFrame and Excel file using pandas
        df = pd.DataFrame(update_info)
        df.to_excel(filename, index=False)

        
    except PermissionError as e:
        print(f"Excel file created: {filename}")

def send_email_with_attachment(recipient, subject, body, attachment_path):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Send()

def create_excel_file(update_info, services_info, filename):
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Write Windows updates information
            df_updates = pd.DataFrame(update_info)
            df_updates.to_excel(writer, sheet_name='Windows Updates', index=False)

            # Write services information
            df_services = pd.DataFrame(services_info)
            df_services.to_excel(writer, sheet_name='Services', index=False)

        print(f"Excel file created: {filename}")
    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Main workflow
update_info = get_windows_updates()
services_info = list_services()

# Define Excel file name based on the current date
excel_file = f"Windows_Updates_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
create_excel_file(update_info, services_info, excel_file)

# Send the Excel file via email
send_email_with_attachment(
    recipient="chris20120330@gmail.com",
    subject="Weekly Windows Update Report",
    body="Please find attached the weekly report of Windows updates and services.",
    attachment_path=win32file.GetFullPathName(excel_file)
)