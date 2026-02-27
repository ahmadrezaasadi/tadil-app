[app]

title = Tadil App
package.name = tadilapp
package.domain = org.tadil

version = 1.0

source.dir = .
source.include_exts = py,png,jpg,kv,atlas

entrypoint = main.py

# کتابخانه‌ها
requirements = python3,kivy==2.2.1,pyjnius==1.6.1,python-bidi==0.4.2,arabic_reshaper,openpyxl,plyer

orientation = portrait

# icon / splash اگر خواستی فعال کن
# android.icon = assets/icon.png
# android.presplash_color = #FFFFFF


[buildozer]

log_level = 2

# 🚀 خیلی مهم
android.skip_update = True

# ✅ نسخه تست شده و پایدار
android.ndk = 27.3.13750724
android.api = 31
android.minapi = 21

# 🔥 فقط یک آرک نگه داریم (مشکل pyjnius رو حل می‌کنه)
android.archs = arm64-v8a

android.bootstrap = sdl2

# نسخه پایدار p4a
p4a.branch = develop

android.add_libs_armeabi_v7a = false
android.add_libs_arm64_v8a = false
