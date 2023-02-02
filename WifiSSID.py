import subprocess
import xlwt
from xlwt import Workbook

data = subprocess.check_output(['netsh', 'wlan', 'show', 'profiles']).decode(
    'utf-8', errors="backslashreplace").split('\n')
profiles = [i.split(":")[1][1:-1] for i in data if "All User Profile" in i]


wifi = Workbook()
sheet1 = wifi.add_sheet('Sheet 1')
a = 0

for i in profiles:
    try:
        results = subprocess.check_output(['netsh', 'wlan', 'show', 'profile', i, 'key=clear']).decode(
            'utf-8', errors="backslashreplace").split('\n')
        results = [b.split(":")[1][1:-1]
                   for b in results if "Key Content" in b]

        try:
            sheet1.write(a, 0, "{0}".format(i, results[0]))
            sheet1.write(a, 1, "{1}".format(i, results[0]))

            print("{:<30} |  {:<}".format(i, results[0]))
            a = a + 1
        except IndexError:
            print("{:<30}|  {:<}".format(i, ""))
    except subprocess.CalledProcessError:
        print("{:<30}|  {:<}".format(i, "ENCODING ERROR"))

wifi.save('WifiSSID.xls')
