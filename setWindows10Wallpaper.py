#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""    DESCRIPTION

        This tool can either set the latest locally stored Microsoft
        Spotlight (--spotlight) image as Desktop wallpaper. Or it fetches
        the latest image from Bing's Image Of The Day (--bing) collection
        and set's it as wallpaper. Or it does randomly one of the two (--random)
        Default: Microsoft Spotlight

    EXAMPLES

        setWindows10Wallpaper.py
        setWindows10Wallpaper.py --spotlight

    REQUIREMENTS

        Requires Python 3.*
        Python PyWin32

    EXIT STATUS

        0: command executed successfully
        2: command failed with errors
"""

import ctypes
import datetime
import glob
import imghdr
import json
import optparse
import os
import random
import requests
import struct
import sys
import win32api
import win32con

__author__ = "Roland Rickborn (gitRigge)"
__copyright__ = "Copyright (C) 2019 Roland Rickborn"
__license__ = "MIT"
__version__ = "0.2"
__status__ = "Development"

def setWallpaperWithCtypes(path):
    """Sets asset given by 'path' as current Desktop wallpaper"""
    cs = ctypes.create_string_buffer(path.encode('utf-8'))
    ok = ctypes.windll.user32.SystemParametersInfoA(win32con.SPI_SETDESKWALLPAPER, 0, cs, 0)

def getLatestWallpaperLocal():
    """Loops through all locally stored Windows Spotlight assets
    and returns the latest asset which has the same orientation as the screen
    """
    list_of_files = sorted(glob.glob(os.environ['LOCALAPPDATA']
        +r'\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets'
        +'\*'), key=os.path.getmtime, reverse=True)
    for asset in list_of_files:
        if imghdr.what(asset) in ['jpeg', 'jpg', 'png', 'gif']:
            if isImageLandscape(asset) == isScreenLandscape():
                return asset
    sys.exit(2)

def get_image_size(fname):
    """Checks if the asset given by 'fname' is of type 'png', 'jpeg' or 'gif',
    reads the dimensions of the asset and returns its width and height in pixels
    """
    with open(fname, 'rb') as fhandle:
        head = fhandle.read(24)
        if len(head) != 24:
            return
        if imghdr.what(fname) == 'png':
            check = struct.unpack('>i', head[4:8])[0]
            if check != 0x0d0a1a0a:
                sys.exit(2)
            width, height = struct.unpack('>ii', head[16:24])
        elif imghdr.what(fname) == 'gif':
            width, height = struct.unpack('<HH', head[6:10])
        elif imghdr.what(fname) == 'jpeg':
            try:
                fhandle.seek(0) # Read 0xff next
                size = 2
                ftype = 0
                while not 0xc0 <= ftype <= 0xcf:
                    fhandle.seek(size, 1)
                    byte = fhandle.read(1)
                    while ord(byte) == 0xff:
                        byte = fhandle.read(1)
                    ftype = ord(byte)
                    size = struct.unpack('>H', fhandle.read(2))[0] - 2
                # We are at a SOFn block
                fhandle.seek(1, 1)  # Skip `precision' byte.
                height, width = struct.unpack('>HH', fhandle.read(4))
            except Exception: #IGNORE:W0703
                sys.exit(2)
        else:
            sys.exit(2)
        return width, height

def isImageLandscape(asset):
    """Checks the orientation of the asset given by 'asset' and returns 'True' if the asset's
    orientation is landscape
    """
    myDim = get_image_size(asset)
    # Calculate Width:Height; > 0 == landscape; < 0 == portrait
    if myDim[0]/myDim[1] > 1:
        return True
    else:
        return False

def isScreenLandscape():
    """Checks the current screen orientation and returns 'True' if the screen orientation
    is landscape
    """
    if getScreenWidth()/getScreenHeight() > 1:
        return True
    else:
        return False

def getScreenWidth():
    """Reads Windows System Metrics and returns screen width in pixel"""
    width = win32api.GetSystemMetrics(0)
    return width

def getScreenHeight():
    """Reads Windows System Metrics and returns screen heigth in pixel"""
    height = win32api.GetSystemMetrics(1)
    return height

def getLatestWallpaperRemote():
    """Retrieves the URL of Bing's Image Of The Day image, downloads the image,
    stores it in a temporary folder and returns the path to it
    """
    # get image url
    response = requests.get("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US")
    image_data = json.loads(response.text)

    image_url = image_data["images"][0]["url"]
    image_url = image_url.split("&")[0]
    full_image_url = "https://www.bing.com" + image_url

    # image's name
    image_name = datetime.date.today().strftime("%Y%m%d")
    image_extension = image_url.split(".")[-1]
    image_name = image_name + "." + image_extension

    # download and save image
    img_data = requests.get(full_image_url).content
    dir_path = os.path.join(os.environ['TEMP'],'BingBackgroundImages')
    os.makedirs(dir_path, exist_ok=True)
    with open(os.path.join(dir_path, image_name), 'wb') as handler:
        handler.write(img_data)
    return os.path.join(dir_path, image_name)

def usage(option, opt, value, parser):
    """Shows help of this tool"""
    myDocstring = ""
    if opt in ["-i", "--info"]:
        myDocstring = myDocstring+"\n"+__doc__
        myDocstring = myDocstring+"\n    AUTHOR\n\n        $author$\n"
        parser.print_help()
    if opt in ["-V", "-i", "--version", "--info"]:
        myDocstring = myDocstring+"\n    VERSION\n\n        $version$"
    myDocstring = myDocstring.replace('$version$', __version__)
    myDocstring = myDocstring.replace('$author$',__author__)
    print(myDocstring)
    sys.exit(0)

if __name__ == "__main__":
    parser = optparse.OptionParser()
    parser.add_option("-b", "--bing", action="store_true", dest="bing", default=False,
        help="set Bing Image Of The Day as wallpaper")
    parser.add_option("-s", "--spotlight", action="store_true", dest="spotlight", default=True,
        help="set Microsoft Spotlight as wallpaper [default]")
    parser.add_option("-r", "--random", action="store_true", dest="random", default=False,
        help="set wallpaper from random source")
    parser.add_option("-V", "--version", action="callback", callback=usage,
        help="show version")
    parser.add_option("-i", "--info", action="callback", callback=usage,
        help="show license and author information")
    (options, args) = parser.parse_args()
    path = ""
    if options.random:
        myChoice = random.choice(['bing', 'spotlight'])
        if myChoice == 'bing':
            options.bing = True
        elif myChoice == 'spotlight':
            options.spotlight = True
    if options.bing:
        print("bing")
        path = getLatestWallpaperRemote()
        setWallpaperWithCtypes(path)
    elif options.spotlight:
        print("spotlight")
        path = getLatestWallpaperLocal()
        setWallpaperWithCtypes(path)
    else:
             print("ERROR: unhandled option")
             usage()
             sys.exit(2)
    sys.exit(0)