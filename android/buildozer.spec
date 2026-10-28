[app]

# Application title
title = Smart PDF Extractor

# Package name
package.name = smartdocextractor

# Package domain
package.domain = org.extractor

# Source code directory
source.dir = .

# Include file extensions
source.include_exts = py,png,jpg,jpeg,kv,atlas,json,txt

# Exclude extensions
source.exclude_exts = spec,pyc,pyo,pyd,bak

# Exclude directories
source.exclude_dirs = tests,bin,.git,.github,__pycache__

# Version
version = 1.0.0

# Application requirements
requirements = python3,kivy,numpy,pillow,plyer,setuptools

# Orientation
orientation = portrait

# Fullscreen
fullscreen = 0

# Presplash (optional)
# presplash.filename = data/presplash.png

# Icon (optional)
# icon.filename = data/icon.png

# Android permissions
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# Android API levels
android.api = 33
android.minapi = 21

# NDK
android.ndk = 25b

# Android architecture
android.archs = arm64-v8a

# Storage mode
android.private_storage = True

# Accept SDK licenses
android.accept_sdk_license = True

# Copy libraries
android.copy_libs = 1

# Enable backups
android.allow_backup = True

# Logcat filtering
android.logcat_filters = *:S python:D

# Generate APK
android.debug_artifact = apk
android.release_artifact = apk

# Keep screen on
android.wakelock = False


[buildozer]

# Buildozer log level
log_level = 2

# Warn when running as root
warn_on_root = 1

# Output directory
bin_dir = ./bin
