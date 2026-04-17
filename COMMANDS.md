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
