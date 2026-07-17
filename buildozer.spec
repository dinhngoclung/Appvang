[app]
title = Tach Vang
package.name = tachvang
package.domain = com.lung.tachvang
source.dir =.
source.include_exts = py,png,jpg,kv,atlas,json
version = 1.0
requirements = python3,kivy,openpyxl
# ĐỪNG thêm pandas vội, ra APK rồi mới thêm. openpyxl đã đủ đọc Excel rồi

orientation = portrait
fullscreen = 0
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.accept_sdk_license_agreement = True
android.ant = auto

[buildozer]
log_level = 2
warn_on_root = 1
