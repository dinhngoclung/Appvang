[app]
title = Tach Vang
package.name = tachvang
package.domain = com.lung.tachvang
source.dir =.
source.include_exts = py,png,jpg,kv,atlas
version = 1.0
requirements = python3,kivy,openpyxl
orientation = portrait
android.permissions = INTERNET,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.accept_sdk_license_agreement = True
android.ant = auto

[buildozer]
log_level = 2
warn_on_root = 1
