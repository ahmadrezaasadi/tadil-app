[app]

title = Tadil App
package.name = tadilapp
package.domain = org.tadil

version = 1.0

source.dir = .
source.include_exts = py,png,jpg,kv,atlas

entrypoint = main.py

# ✅ کتابخانه‌های مورد نیاز
requirements = python3,kivy==2.2.1,pyjnius==1.6.1,python-bidi==0.4.2,arabic_reshaper,openpyxl,plyer

orientation = portrait

# android.icon = assets/icon.png
# android.presplash_color = #FFFFFF


[buildozer]

log_level = 2

# ✅ جلوگیری از آپدیت بی‌خودی SDK هر بار
android.skip_update = True

# ✅ نسخه پایدار NDK
android.ndk = 23.2.8568313

android.api = 33
android.minapi = 21

android.archs = arm64-v8a, armeabi-v7a

android.bootstrap = sdl2

# ✅ نسخه پایدار python-for-android
p4a.branch = master

android.add_libs_armeabi_v7a = false
android.add_libs_arm64_v8a = false
