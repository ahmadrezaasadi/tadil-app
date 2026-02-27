[app]

# (اسم اپلیکیشن تو)
title = Tadil App

# بسته اندروید
package.name = tadilapp
package.domain = org.tadil

# نسخه اپ
version = 1.0

source.dir = .

# فایل اصلی اجرای اپ
source.include_exts = py,png,jpg,kv,atlas

# مین فایل اجرا
entrypoint = main.py

# نیازمندی های پایتون
requirements = python3,kivy==2.2.1,python-bidi==0.4.2,arabic_reshaper,openpyxl,plyer

orientation = portrait

# آیکون اپ
# android.icon = assets/icon.png

# اسپلش اسکرین
# android.presplash_color = #FFFFFF

[buildozer]

log_level = 2

# 🔥 جلوگیری از آپدیت بیخود هر بار
android.skip_update = True

# 🔥 نسخه پایدار NDK
android.ndk = 23.2.8568313

# اندروید API
android.api = 33
android.minapi = 21

# معماری ها
android.archs = arm64-v8a, armeabi-v7a

# bootstrap پایدار
android.bootstrap = sdl2

# 🔥 نسخه پایدار p4a
p4a.branch = develop

# استفاده از لایبرری ها
android.add_libs_armeabi_v7a = false
android.add_libs_arm64_v8a = false
