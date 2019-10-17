@echo off
schtasks /Create /SC HOURLY /TN WallpaperTask /TR "python setWindows10Wallpaper.py"