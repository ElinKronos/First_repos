
from datetime import datetime

print("Hello world!")

now = datetime.now()

format_date = now.strftime("%Y-%m-%d %H:%M:%S")
print(format_date)

format_date_only = now.strftime("%A, %d %B %Y")
print(format_date_only)

format_time_only = now.strftime("%I:%M %p")
print(format_time_only)

format_date_2 = now.strftime("%d.%m.%Y")
print(format_date_2)

