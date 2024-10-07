### Project Title

Email Reader

### Description

This project is a console application written in C# that reads emails from a specified mailbox and folder using Microsoft Graph API. The application authenticates using Azure Identity and retrieves messages based on the provided command-line arguments.

### Features

- Authenticate using Azure Device Code Flow
- Read emails from a specified mailbox and folder
- Display email details such as sender, recipients, subject, and body

### Prerequisites

- .NET SDK
- Azure subscription and Azure AD app registration
- Microsoft Graph API permissions: `User.Read` and `Mail.Read`

### Setup

1. **Clone the repository:**
   ```sh
   git clone <repository-url>
   cd <repository-directory>
   ```

2. **Set up Azure AD app registration:**
   - Register an app in Azure AD and note down the `Client ID` and `Tenant ID`.
   - Grant the app `User.Read` and `Mail.Read` permissions.

3. **Set environment variables:**
   ```sh
   export CLIENT_ID=<your-client-id>
   export TENANT_ID=<your-tenant-id>
   ```

### Usage

Run the application with the following command:
```sh
 dotnet run -- --mailbox <mailbox-address> --folder <folder-name> --messages <message-count>
```

#### Command-Line Arguments
- `--mailbox`: The email address of the mailbox to read from (default: "me" for the authenticated user)
- `--folder`: The name of the folder to read emails from (default: "Inbox")
- `--messages`: The number of messages to read (default: 10)

### Example

```sh
 dotnet run -- --mailbox user@example.com --folder Inbox --messages 5
```

### Code Overview

- **Authentication**: The application uses `DeviceCodeCredential` for authentication, which prompts the user to authenticate using a device code.
- **Graph Client**: An instance of `GraphServiceClient` is created using the authenticated credential.
- **Email Retrieval**: Emails are fetched from the specified folder and details such as sender, recipients, subject, and body are displayed.

### Error Handling

- If `CLIENT_ID` or `TENANT_ID` environment variables are not set, the application will exit with code 3.
- If invalid command-line arguments are provided, the application will display an error message and exit with code 1.
- If the folder is not found, the application will display an error message and exit with code 2.

### Dependencies

- Microsoft.Graph
- Azure.Identity
- Microsoft.Identity.Client

### License

This project is licensed under the MIT License.

### Acknowledgements

- [Microsoft Graph](https://docs.microsoft.com/en-us/graph/overview)
- [Azure Identity](https://docs.microsoft.com/en-us/dotnet/api/overview/azure/identity-readme)

