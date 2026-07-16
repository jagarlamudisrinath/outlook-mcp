# Outlook MCP Server (win32com)

An MCP server that controls the **Microsoft Outlook desktop app** through COM
automation (`win32com`). It does **not** use Microsoft Graph — there is no
Azure app registration, no OAuth, no API keys. It operates on whatever
account(s) your local Outlook client is already signed into, which makes it
ideal for corporate environments where Graph API access is locked down.

## Requirements

- **Windows** (COM automation is Windows-only)
- **Classic Outlook desktop** installed and configured with an account.
  The "new Outlook" (the web-based rewrite) does **not** expose the COM
  object model and will not work — switch the toggle back to classic Outlook.
- Python 3.10+

## Installation

```powershell
git clone <this-repo>
cd outlook_mcp
pip install -e .
```

This installs `mcp` and `pywin32` and registers the `outlook-mcp` command.

## Configure Claude Desktop / Claude Code

Claude Desktop (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp"
    }
  }
}
```

Claude Code:

```powershell
claude mcp add outlook -- outlook-mcp
```

If Outlook is not running, the first tool call starts it automatically.

## Tools

| Tool | Description |
|------|-------------|
| `list_folders` | Tree of all mail folders across every account/store |
| `list_emails` | Recent emails in a folder (newest first, optional unread-only) |
| `get_email` | Full email by `entry_id`, including attachment list |
| `search_emails` | Search a folder by subject / sender / body text |
| `save_attachments` | Save an email's attachments to a local directory |
| `send_email` | Send (or draft) a new email — plain text or HTML, with attachments |
| `reply_to_email` | Reply / reply-all, with the original quoted below |
| `forward_email` | Forward an email, attachments included |
| `mark_email` | Mark read / unread |
| `move_email` | Move an email to another folder |
| `delete_email` | Move an email to Deleted Items |
| `list_calendar_events` | Events in a window around today (recurrences expanded) |
| `create_calendar_event` | Create an appointment or send a meeting invite |
| `check_availability` | Show attendees' free/busy timeline for a day |
| `find_meeting_times` | Find slots when **all** attendees are free |
| `schedule_out_of_office` | Block your calendar with Out-of-Office status |
| `set_automatic_replies` | Guidance only — auto-replies aren't COM-settable (see below) |
| `list_contacts` | List/search the default Contacts folder |

Emails are addressed by their Outlook `EntryID`, which `list_emails` and
`search_emails` return — pass it to `get_email`, `reply_to_email`, etc.
Folders accept well-known names (`Inbox`, `Sent Items`, `Drafts`, ...) or
slash paths like `Inbox/Receipts` or `you@company.com/Inbox` to target a
specific account.

## Example prompts

- "Show my unread emails"
- "Search my inbox for emails from Alice about the Q3 budget"
- "Reply to that email saying I'll review it by Friday"
- "Save the attachments from that email to C:\\Users\\me\\Downloads"
- "What's on my calendar this week?"
- "Check if alice@corp.com and bob@corp.com are free Thursday afternoon"
- "Find a 30-minute slot next Monday when the whole team is available"
- "Set up a 30-minute meeting with bob@example.com tomorrow at 2pm"
- "Mark me out of office next Friday through the following Wednesday"

## Checking other people's availability

`check_availability` and `find_meeting_times` read **free/busy** data — the
same busy/free blocks Outlook shows in the Scheduling Assistant. In most
Exchange/Microsoft 365 organizations every user can see everyone else's
free/busy by default (busy times only, not the meeting subjects), so no
special mailbox permissions are needed. If a person has restricted their
free/busy sharing, their slots come back as unavailable and the tools say so.

## Out of office: what works and what doesn't

- **`schedule_out_of_office` works** — it creates a calendar block with
  "Out of Office" availability status, so you show as away in other people's
  free/busy and Scheduling Assistant.
- **Automatic replies (the auto-responder email) cannot be set via COM.** The
  Outlook Object Model simply does not expose Out-of-Office reply settings.
  `set_automatic_replies` therefore only returns guidance. To turn on the
  auto-responder, either flip it on manually (File > Automatic Replies) or use
  Exchange Web Services (`SetUserOofSettings`) / Microsoft Graph
  (`mailboxSettings.automaticRepliesSetting`) — both of which are outside COM.

## Notes & troubleshooting

- **Security prompts:** depending on your organization's policy, Outlook may
  show an "Allow access?" dialog when a program reads addresses or sends
  mail programmatically. Your admin controls this via Group Policy
  ("Programmatic Access" settings in Trust Center).
- **"New Outlook" toggle:** if tool calls fail with "class not registered" or
  Outlook opens but nothing happens, you are likely on new Outlook. Switch
  back to classic Outlook.
- **Bitness/permissions:** run the MCP server as the same user (and not
  elevated differently) as Outlook, or COM will refuse to connect to the
  running instance.
