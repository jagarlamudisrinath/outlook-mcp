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

**Mail — read & search**

| Tool | Description |
|------|-------------|
| `list_folders` | Tree of all mail folders across every account/store |
| `list_emails` | Recent emails in a folder (newest first, optional unread-only) |
| `get_email` | Full email by `entry_id` — SMTP addresses, headers, HTML, links, Message-ID |
| `search_emails` | Filtered search: text, sender/recipient, date range, unread/flagged/importance/attachments, across all folders or subfolders |
| `list_attachments` | List an email's attachments (index, size, inline vs. real, MIME) |
| `save_attachments` | Save attachments safely (sanitized names, no collisions, size limit, SHA-256) |

**Mail — write, drafts & management**

| Tool | Description |
|------|-------------|
| `send_email` | Send/draft — HTML, attachments, importance/sensitivity/categories, send-as/on-behalf, scheduled delivery, read receipt |
| `reply_to_email` | Reply / reply-all — HTML, attachments |
| `forward_email` | Forward — CC/BCC, validated recipients, attachments |
| `list_drafts` | List messages in the Drafts folder |
| `update_draft` | Edit a saved draft (subject/body/recipients/attachments) |
| `send_draft` | Send an existing draft (re-validates recipients) |
| `delete_draft` | Delete a draft |
| `remove_attachment` | Remove one attachment from a draft/item |
| `mark_email` | Mark read / unread |
| `move_email` | Move an email to another folder |
| `delete_email` | Move an email to Deleted Items |

**Calendar & scheduling**

| Tool | Description |
|------|-------------|
| `list_calendar_events` | Events in a window around today (recurrences expanded) |
| `get_calendar_event` | Full event details — attendees + responses, recurrence, join links, categories |
| `search_calendar_events` | Filter events by subject/organizer/location/category/Teams/recurring/attachments |
| `create_calendar_event` | Create an appointment or send a meeting invite (timezone-aware) |
| `create_teams_meeting` | Send a meeting invite (Teams link via org setting, see below) |
| `create_recurring_meeting` | Send a recurring series (daily/weekly/monthly/yearly, timezone-aware) |
| `update_calendar_event` | Update a whole event/series; add/remove attendees, optionally re-invite |
| `update_single_occurrence` | Edit one occurrence of a recurring series |
| `cancel_calendar_event` | Cancel a meeting you organize (notifies attendees) |
| `cancel_single_occurrence` | Cancel one occurrence of a series |
| `delete_calendar_event` | Delete an event with no notice |
| `check_availability` | Show attendees' free/busy timeline for a day |
| `find_meeting_times` | Find slots when **all** attendees are free, across a date range |
| `schedule_out_of_office` | Block your calendar with Out-of-Office status |
| `set_automatic_replies` | Guidance only — auto-replies aren't COM-settable (see below) |

**Meeting invitations, contacts & accounts**

| Tool | Description |
|------|-------------|
| `list_meeting_invitations` | Meeting requests in your Inbox — response state + calendar conflicts |
| `respond_to_invitation` | Accept / tentative / decline an invitation |
| `list_shared_calendar` | Read another person's calendar (if shared to you) |
| `list_contacts` | List/search the default Contacts folder |
| `list_accounts` | List sending accounts and mail stores in the profile |
| `outlook_health` | Diagnostics: Outlook version, user, store, connection mode |

Emails and events are addressed by their Outlook `EntryID`, which the list/
search tools return — pass it to `get_email`, `reply_to_email`,
`update_calendar_event`, etc. Folders accept well-known names (`Inbox`,
`Sent Items`, `Drafts`, ...) or slash paths like `Inbox/Receipts` or
`you@company.com/Inbox` to target a specific account.

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

## Time zones (recommended for meetings)

The meeting-creation tools (`create_calendar_event`, `create_teams_meeting`,
`create_recurring_meeting`) accept an optional `timezone_name`. Give it an IANA
name (`Asia/Kolkata`), a Windows ID (`India Standard Time`), or a common
abbreviation (`IST`). The start time you pass is treated as wall-clock time in
that zone, and the result reports both the local and UTC times.

This matters most for **recurring** meetings: without an explicit zone, a
series pinned to "12:00" can shift by an hour when it crosses a daylight-saving
boundary, because Outlook interprets the time in whatever zone the profile
happens to be in. Passing `timezone_name` sets the appointment's
`StartTimeZone`/`EndTimeZone` in COM so occurrences stay put. (Note: on Windows,
Python needs the `tzdata` package for this — it's declared as a dependency.)

## Recipient validation

`send_email` (and the meeting tools) add each recipient individually and call
`ResolveAll()` before sending. If any name is ambiguous or unknown, the send is
**rejected** with the offending entries listed — pass `allow_unresolved=True`
to override. Results list every recipient's canonical SMTP address (Exchange
`EX` addresses are resolved to real SMTP), so you can confirm exactly who will
receive the message before trusting it.

## Development & tests

```powershell
pip install -e ".[dev]"
pytest
```

The pure-logic layer (date/timezone parsing, DST conversion, recurrence,
recipient resolution) is covered by mock-based tests that run on any platform —
no Windows or Outlook required. COM interaction itself can only be exercised on
Windows. CI runs the suite on Python 3.10–3.12.

## Checking other people's availability

`check_availability` and `find_meeting_times` read **free/busy** data — the
same busy/free blocks Outlook shows in the Scheduling Assistant. In most
Exchange/Microsoft 365 organizations every user can see everyone else's
free/busy by default (busy times only, not the meeting subjects), so no
special mailbox permissions are needed. If a person has restricted their
free/busy sharing, their slots come back as unavailable and the tools say so.

## Teams meetings and invitations

- **Responding to invites works fully.** `list_meeting_invitations` shows
  pending requests in your Inbox; `respond_to_invitation` accepts, tentatively
  accepts, or declines them (with an optional note, and an option to respond
  without notifying the organizer).
- **Reading others' meetings** (`list_shared_calendar`) works when that person
  has shared their calendar with you at Reviewer permission or higher. It shows
  real subjects/times, unlike `check_availability`, which shows only free/busy.
- **Creating a Teams meeting is partial.** COM **cannot** inject a Teams join
  link — the link is produced by the Teams Meeting Add-in/service, not the
  Outlook Object Model. `create_teams_meeting` sends the invite and relies on
  the mailbox setting **File > Options > Calendar > "Add online meeting to all
  meetings"**; with that ON, sent meetings automatically become Teams meetings.
  With it OFF, an ordinary meeting invite is sent and you'd add the Teams link
  manually.
- **"Propose new time" is not available via COM.** The Object Model has no
  propose-new-time method. To suggest another slot, decline with a message (or
  use `find_meeting_times` to pick a slot and send a fresh invite).
- **Recurring meetings are supported.** `create_recurring_meeting` sends a full
  series — daily, weekly (with specific weekdays like `Mon,Wed,Fri`), monthly,
  or yearly, with an `interval` (e.g. every 2 weeks) and an end defined by a
  number of occurrences (`count`) or an end date (`until`). Attendees receive
  the series and accept it as a series. Recurring invites you *receive* are
  accepted/declined as a whole series by `respond_to_invitation`.

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
