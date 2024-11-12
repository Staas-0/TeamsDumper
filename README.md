# Microsoft Teams Dumper

The repository contains a PowerShell script that allows you to export your Microsoft Teams chat conversations, in HTML format, to your local disk.


## Getting Started

### Prerequisites

- An Azure AD app registration needs to be created with the following permissions:
  -  User chats: Chat.Read.All, ChatMember.Read.All, ChatMessage.Read.All, User.Read.All

  - Teams channels: Team.ReadBasic.All, TeamMember.Read.All, Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All, Files.Read.All, OnlineMeetingRecording.Read.All, OnlineMeetingTranscript.Read.All


- Follow the steps at https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app and https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-configure-app-access-web-apis

- You must be running PowerShell 7. See https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-windows?view=powershell-7

### Steps

1. Download this repository

1. Create a folder where you'll have your chat history exported to, otherwise it will be exported to \Exports from where the script is ran

1. Run the Powershell script

   ```PowerShell
   PS> Get-Help ./TeamsDumper.ps1
   ...
   PS> ./TeamsDumper.ps1 -ExportFolder C:\Users\<you>\OneDrive\ExportChat
   ```

1. Watch the slow crawl magic happen, exporting your chat history.

### Options

- \-verbose for more detailed output/debugging
- \-exportFolder to specify output directory
- clientId, clientSecret, and tenantId can all be set with arguments, but the script will also prompt for them if not set.

### Improvement Ideas

- Grab files attachments from SharePoint links
- ...?

#### Credits
Inspired by
https://github.com/telstrapurple/MSTeamsChatExporter and https://github.com/evenevan/export-ms-teams-chats