# -Microsoft-graph-API-Script_to_Automated_Email-_to_remind_users_Passwords_Expiracy_English_v.
# Password Expiration Notification Script

This script is designed to send email notifications to users, informing them when their Active Directory (AD) passwords are about to expire. It connects to Microsoft Graph API to send emails and retrieves user data from Active Directory using the `ldap3` library. This allows IT administrators to automate password expiration reminders for users.

## Features
- **Automated Email Notifications**: Sends password expiration reminders to users whose passwords are set to expire within a specified number of days.
- **Active Directory Integration**: Connects to AD to retrieve user information, including email addresses and password last set dates.
- **Logging**: Logs activities and events, including API responses and any errors, to a CSV file for easy tracking.
- **Testing Mode**: Includes a testing mode that allows sending notifications to a specific test email rather than all users.

## Requirements
- **Dependencies**: The script requires the following Python packages:
  - `msal`: For Microsoft Authentication Library, enabling access to Microsoft Graph API.
  - `requests`: For making HTTP requests to the Microsoft Graph API.
  - `ldap3`: For interacting with Active Directory.
- **Permissions**: Microsoft Graph API permissions to send emails on behalf of a user, typically configured in Azure AD for your application.

## Configuration

1. **Microsoft Graph API**: Register an application in Azure AD to get the necessary credentials:
   - **Client ID**: Application ID of your Azure AD app.
   - **Tenant ID**: Directory ID for your Azure AD.
   - **Client Secret**: Secret key generated for your Azure AD app.

2. **Active Directory Account**: Configure an account with appropriate permissions to read user information from AD.

3. **Environment Variables**: Replace the placeholders with your actual configuration:
   - `client_id`: Your Azure AD Application (Client) ID.
   - `tenant_id`: Your Azure AD Directory (Tenant) ID.
   - `client_secret`: The client secret for your Azure AD app.
   - `user_email`: Email address from which notifications will be sent.

4. **Settings for Notification**:
   - `expire_in_days`: Number of days before password expiration when users should receive a notification.
   - `testing_enabled`: Boolean flag to toggle test mode. When enabled, notifications are sent to a single test email.
   - `test_recipient`: Email address to receive test notifications when `testing_enabled` is `True`.
   - `logging_enabled`: Boolean flag to enable logging. Logs are stored in a CSV file generated during script execution.

