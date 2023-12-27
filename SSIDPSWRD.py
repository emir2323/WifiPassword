import subprocess
import openpyxl


result = subprocess.run(["netsh", "wlan", "show", "profiles"], capture_output=True, text=True)


if result.stdout is None:
    print("Wi-Fi ağlarına erişilemiyor. Lütfen komutu yönetici olarak çalıştırdığınızdan emin olun.")
else:
    profiles = [line.split(":")[1].strip() for line in result.stdout.splitlines() if "All User Profile" in line]

   
    wifi_data = []
    for profile in profiles:
        result = subprocess.run(["netsh", "wlan", "show", "profile", profile, "key=clear"], capture_output=True, text=True)
        ssid = profile.strip()
        
       
        if result.stdout:
            password_line = next((line for line in result.stdout.splitlines() if "Key Content" in line), None)
        else:
            password_line = None
        
        
        if password_line:
            password = password_line.split(":")[1].strip()
        else:
            password = "Şifre bulunamadı"
        
        wifi_data.append((ssid, password))

   
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Wi-Fi Bilgileri"

   
    sheet.append(["SSID", "Şifre"])

   
    for ssid, password in wifi_data:
        sheet.append([ssid, password])

   
    workbook.save("wifi_bilgileri.xlsx")

    print("Wi-Fi bilgileri başarıyla kaydedildi.")



