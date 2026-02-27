[app]
title = Tadil App
package.name = tadilapp
package.domain = org.tadil

version = 1.0

source.dir = .
source.include_exts = py,png,jpg,kv,atlas

entrypoint = main.py

requirements = python3,kivy==2.2.1,python-bidi==0.4.2,arabic_reshaper,openpyxl,plyer

orientation = portrait

[buildozer]
log_level = 2

android.api = 31
android.minapi = 21

android.ndk = 25.2.9519653
android.sdk = 33

android.archs = arm64-v8a, armeabi-v7a

android.bootstrap = sdl2

p4a.branch = develop

android.skip_update = True
