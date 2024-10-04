# o365_imap_migration

Steps:

1. Install all the requirements.
2. Register a new app in the O365 Entra Admin Center.

3. Grant Microsoft Graph API permissions (ensure you select "Application permissions," not "Delegated permissions") and provide admin consent for the following     
   permissions:

    Mail.Read, 
    Mail.ReadWrite, 
    User.Read.All


4. Go to "Certificates & Secrets" → "Client Secrets," and create a new client secret.
5. Save the following information:

   Application (client) ID, 
   Directory (tenant) ID, 
   Client secret value
   
6. Add all of this information to the config file. Also, select in the config file if you want to migrate attachments and choose the email format: 'html' or 'plain' (note that some mail providers do not support HTML format emails).

7. Populate your CSV file with user data and IMAP server details as shown in the example. Run the code, and it will migrate all the emails.
