from datetime import datetime, timedelta

today = datetime.now()
days_ahead = (5 - today.weekday()) % 7 or 7
next_saturday = today + timedelta(days=days_ahead)
next_saturday = next_saturday.replace(hour=3, minute=30, second=0, microsecond=0)

print(next_saturday.strftime('%Y-%m-%dT%H:%M:%S'))

