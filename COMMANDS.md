# Commands Cheat Sheet

## Matrix Commands

Show help:

```text
!help
```

Start today's entry:

```text
!today
```

Show today's status:

```text
!status
```

Show a specific date:

```text
!status 17.04
```

Correct a normal day:

```text
!correct 17.04 08:30 17:10 12:00-12:30
```

Correct a special day:

```text
!correct 17.04 k
!correct 17.04 u
!correct 17.04 g
```

Show or change the response language:

```text
!language
!language de
!language en
```

Trigger a test reminder:

```text
!testreminder
```

## Service Commands

Check if the bot is running:

```bash
systemctl --user status work-hours-bot.service
```

Short status only:

```bash
systemctl --user is-active work-hours-bot.service
```

Start:

```bash
systemctl --user start work-hours-bot.service
```

Stop:

```bash
systemctl --user stop work-hours-bot.service
```

Restart:

```bash
systemctl --user restart work-hours-bot.service
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

Run without systemd:

```bash
bash run.sh
```

## SAP Zeitnachweis Helper

Install the browser automation dependency once:

```bash
python -m playwright install firefox
```

Preview the values that would be entered for a month:

```bash
python -m src.sap.fill_zeitnachweis --month 4 --dry-run
```

For Windows Firefox with DLR single sign-on, copy a browser-side filler script to the Windows clipboard:

```bash
python -m src.sap.fill_zeitnachweis --month 4 --browser-js --copy
```

Then open the SAP Zeitnachweis form in your normal Windows Firefox, press `F12`, open the Console, paste the script, and press Enter. The script fills fields in the already-authenticated page and shows an alert. Review everything before saving.

Open SAP in an automation Firefox profile and fill the form after manual login/review:

```bash
python -m src.sap.fill_zeitnachweis --month 4
```

The helper fills `2003568` into section II row 0, writes normal workday values to `ZN_KTR.0.KTRDD`, and writes `U`/`K` days to the absence fields `ABDD`. It pauses before filling and again before closing the browser; it does not submit the form.

## Useful Test Flow

1. Watch logs:

```bash
journalctl --user -u work-hours-bot.service -f
```

2. In Matrix send:

```text
!testreminder
```

3. Reply with:

```text
yes
09:00
17:30
12:00-12:30
```
