[app]
title = Tach Vang
package.name = tachvang
package.domain = com.lung.tachvang
source.dir =.
source.include_exts = py,png,jpg,kv,atlas
version = 1.0
requirements = python3,kivy,openpyxl
orientation = portrait
fullscreen = 0
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.accept_sdk_license_agreement = True
p4a.branch = master
p4a.fork = kivy

[buildozer]
log_level = 2
warn_on_root = 1
