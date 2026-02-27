[app]

title = Tadil App
package.name = tadilapp
package.domain = org.tadil

source.dir = .
source.include_exts = py,ttf
source.include_patterns = assets/*,*.ttf

version = 1.0

requirements = python3,kivy==2.2.1,arabic_reshaper,python-bidi==0.4.2,openpyxl,plyer

orientation = portrait

fullscreen = 0

android.api = 33
android.minapi = 21
android.sdk = 33
android.ndk = 25b

android.archs = arm64-v8a, armeabi-v7a

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

android.allow_backup = True

# برای جلوگیری از حذف openpyxl
android.add_libs_armeabi_v7a =
android.add_libs_arm64_v8a =

[buildozer]

log_level = 2

warn_on_root = 1
