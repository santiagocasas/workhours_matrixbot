# Work Hours Bot

This bot asks for your work times in Matrix, writes them into an Excel timesheet, and can optionally copy the updated workbook to a second path such as a Windows-mounted location.

## What It Does

- Sends a weekday prompt at `16:30`
- Sends a retry reminder at `09:00` the next weekday if the previous day is still pending
- Checks hourly for missed prompts after sleep/resume or downtime
- Writes work start/end times and up to 3 break ranges into an Excel workbook
- Writes special day codes:
  - `K` = sick day
  - `U` = vacation day
  - `G` = flex day
- Skips weekends, holidays, and bridge days automatically when the workbook marks them
- Shows a preview of the expected recorded hours
- Can copy the updated workbook to a second path automatically after each successful write
- Warns you when a month looks complete so you can pull the workbook back from the VPS
- Runs well as a `systemd --user` service in WSL

## Public-Safe Files

Do not commit your real `config.yaml`.

- `config.yaml` is ignored by git
- `config.yaml.example` is the safe template to copy and fill in locally
- `data/state.json` is ignored by git

Create your local config with:

```bash
cp config.yaml.example config.yaml
```

Then edit `config.yaml` with your own Matrix IDs, password, room ID, and file paths.

## Config

Example structure:

```yaml
matrix:
  homeserver: "https://matrix.org"
  bot_user_id: "@yourbot:matrix.org"
  bot_password: "your-bot-password"
  room_id: "!yourRoomId:matrix.org"
  allowed_user_id: "@you:matrix.org"

excel:
  path: "/path/to/your/timesheet.xlsx"
  windows_path: "/mnt/c/Users/youruser/Documents/timesheet.xlsx"

timezone: "Europe/Berlin"

schedule:
  daily_prompt: "16:30"
  morning_retry: "09:00"
  catchup_interval_minutes: 60

state_file: "/path/to/work-hours-bot/data/state.json"
```

`windows_path` is optional. If set, the bot copies the workbook there after each successful write.
`catchup_interval_minutes` is optional and defaults to `60`. It lets the bot recover prompts missed while the machine, WSL session, or service was unavailable at the exact scheduled time.

If you run the bot on a VPS and keep the workbook there, the bot will tell you when a month looks complete. At that point you can pull the workbook back locally with `scripts/pull_workbook_from_vps.sh`.

## Matrix Setup

The bot uses a separate Matrix account and listens only in one room.

Recommended setup:

- Create a dedicated bot account
- Create a private room
- Invite the bot account to that room
- Keep room encryption off for this bot
- Put the bot user ID, allowed user ID, room ID, and password in `config.yaml`

## Commands

Commands are always in English.

- `!help`
  - show command help in the current response language
- `!today`
  - start today's entry if today is not already filled
- `!status`
  - show today's status
- `!status 17.04`
  - show status for a specific date
- `!correct 17.04 u`
  - set a special code for a specific date
- `!correct 17.04 08:30 17:10 12:00-12:30`
  - overwrite a day with times and breaks
- `!language`
  - show the current response language
- `!language de`
  - switch bot replies to German
- `!language en`
  - switch bot replies to English
- `!testreminder`
  - trigger the daily reminder immediately for testing

When all required workdays for a month are filled, the bot also adds a note telling you to pull the workbook back from the VPS.

## Conversation Flow

The bot prompts once per weekday and asks whether you worked that day.

Accepted answers:

- `yes` or `ja`
  - continue with start time, end time, and breaks
- `k`
  - mark the day as sick
- `u`
  - mark the day as vacation
- `g`
  - mark the day as flex day
- `skip`
  - leave the day untouched

Example:

```text
!today
yes
09:00
17:30
12:00-12:30
```

## Systemd Service

Example user service path:

- `~/.config/systemd/user/work-hours-bot.service`

Useful commands:

```bash
systemctl --user status work-hours-bot.service
systemctl --user start work-hours-bot.service
systemctl --user stop work-hours-bot.service
systemctl --user restart work-hours-bot.service
systemctl --user enable work-hours-bot.service
```

## Logs

Show recent logs:

```bash
journalctl --user -u work-hours-bot.service -n 50 --no-pager
```

Watch logs live:

```bash
journalctl --user -u work-hours-bot.service -f
```

## Manual Run

```bash
bash run.sh
```

## Pull Workbook From VPS

Run this locally when you want to copy the current workbook back from the VPS into WSL and Windows:

```bash
bash scripts/pull_workbook_from_vps.sh
```

Optional arguments:

- arg 1: SSH host alias, default `ionosvps`
- arg 2: remote workbook path, default `/home/santi/Personal/matrix-instances/work-hours-bot/data/Arbeitszeitkarte 2026.xlsx`
- arg 3: local WSL target path, default `$HOME/RDM-Software/arbeitszeit/Arbeitszeitkarte 2026.xlsx`
- arg 4: Windows-mounted target path, default `/mnt/c/Users/casa_sa/Documents/Admin/Arbeitszeit/2026/Arbeitszeitkarte 2026.xlsx`

## Files

- `config.yaml.example`
  - safe config template for public repos
- `config.yaml`
  - local runtime configuration, ignored by git
- `run.sh`
  - manual launcher
- `scripts/pull_workbook_from_vps.sh`
  - pulls the current workbook from the VPS back into WSL and Windows
- `src/__main__.py`
  - app entry point
- `src/bot/conversation.py`
  - Matrix commands, reply language handling, and message flow
- `src/bot/matrix_client.py`
  - Matrix login and sync loop
- `src/excel/handler.py`
  - Excel read/write logic and optional second-path copy
- `src/excel/time_utils.py`
  - parsing and preview-hours calculation
- `src/state.py`
  - persistent JSON state
- `src/scheduler.py`
  - scheduled prompts

## Notes

- Excel formulas are recalculated by Excel when you open the workbook
- The bot writes only the input cells and computes a preview itself
- If your workbook layout differs, adjust `src/excel/handler.py`
