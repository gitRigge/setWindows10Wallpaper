#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""    DESCRIPTION

        This tool can either set the latest locally stored Microsoft
        Spotlight (--spotlight) image as Desktop wallpaper. Or it fetches
        the latest image from Bing's Image Of The Day (--bing) collection
        and set's it as wallpaper. Or it does randomly one of the two (--random)
        Default: Microsoft Spotlight

    EXAMPLES

        setWindows10Wallpaper_cli.py
        setWindows10Wallpaper_cli.py --spotlight

    REQUIREMENTS

        Requires Python 3.*
        Python PyWin32

    EXIT STATUS

        0: command executed successfully
        2: command failed with errors
"""

import argparse
import ctypes
import datetime
import glob
import imghdr
import json
import logging
import os
import random
import re
import shutil
import sqlite3
import struct
import sys

import requests
import win32api
import win32con

__author__ = "Roland Rickborn (gitRigge)"
__copyright__ = "Copyright (C) 2020 Roland Rickborn"
__license__ = "MIT License (see https://en.wikipedia.org/wiki/MIT_License)"
__version__ = "0.7"
__status__ = "Development"

def setWallpaperWithCtypes(path):
    """Sets asset given by 'path' as current Desktop wallpaper"""

    logging.debug('setWallpaperWithCtypes({})'.format(path))

    cs = ctypes.create_string_buffer(path.encode('utf-8'))
    ok = ctypes.windll.user32.SystemParametersInfoA(win32con.SPI_SETDESKWALLPAPER, 0, cs, 0)

def getLatestWallpaperLocal():
    """Loops through all locally stored Windows Spotlight assets
    and copies and returns the latest asset which has the same orientation as the screen
    """

    logging.debug('getLatestWallpaperLocal()')

    list_of_files = sorted(glob.glob(os.environ['LOCALAPPDATA']
        +r'\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets'
        +'\*'), key=os.path.getmtime, reverse=True)
    for asset in list_of_files:
        extension = imghdr.what(asset)
        if extension in ['jpeg', 'jpg', 'png']:
            if isImageLandscape(asset) == isScreenLandscape():
                # Generate pseudo url
                full_image_url = os.path.split(asset)[1]
                if not existsImageInDatabase(full_image_url):
                    image_name = getGeneratedImageName(asset+'.'+extension)
                    addImageToDatabase(full_image_url, image_name, "spotlight")
                    dir_path = os.path.join(os.environ['TEMP'],'WarietyWallpaperImages')
                    os.makedirs(dir_path, exist_ok=True)
                    full_image_path = os.path.join(dir_path, image_name)
                    shutil.copyfile(asset, full_image_path)
                    updateImageInDatabase(full_image_url, full_image_path)
                    logging.debug('getLatestWallpaperLocal - full_image_path = {}'.format(full_image_path))
                    return full_image_path
                else:
                    logging.debug('getLatestWallpaperLocal - getImagePathFromDatabase({})'.format(full_image_url))
                    return getImagePathFromDatabase(full_image_url)

def get_image_size(fname):
    """Checks if the asset given by 'fname' is of type 'png', 'jpeg' or 'gif',
    reads the dimensions of the asset and returns its width and height in pixels
    """
    
    logging.debug('get_image_size({})'.format(fname))

    with open(fname, 'rb') as fhandle:
        head = fhandle.read(24)
        if len(head) != 24:
            return
        if imghdr.what(fname) == 'png':
            check = struct.unpack('>i', head[4:8])[0]
            if check != 0x0d0a1a0a:
                logging.debug('get_image_size - Stopping application with exit code "2"\n')
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
                logging.debug('get_image_size - Stopping application with exit code "2"\n')
                sys.exit(2)
        else:
            logging.debug('get_image_size - Stopping application with exit code "2"\n')
            sys.exit(2)
        logging.debug('get_image_size - width, height = {}, {}'.format(width, height))
        return width, height

def isImageLandscape(asset):
    """Checks the orientation of the asset given by 'asset' and returns 'True' if the asset's
    orientation is landscape
    """
    
    logging.debug('isImageLandscape({})'.format(asset))

    myDim = get_image_size(asset)
    # Calculate Width:Height; > 0 == landscape; < 0 == portrait
    if myDim[0]/myDim[1] > 1:
        logging.debug('isImageLandscape - True')
        return True
    else:
        logging.debug('isImageLandscape - False')
        return False

def isScreenLandscape():
    """Checks the current screen orientation and returns 'True' if the screen orientation
    is landscape
    """

    logging.debug('isScreenLandscape()')

    if getScreenWidth()/getScreenHeight() > 1:
        logging.debug('isImageLandscape - True')
        return True
    else:
        logging.debug('isImageLandscape - False')
        return False

def getScreenWidth():
    """Reads Windows System Metrics and returns screen width in pixel"""

    logging.debug('getScreenWidth()')

    width = win32api.GetSystemMetrics(0)
    logging.debug('getScreenWidth - width = {}'.format(width))
    return width

def getScreenHeight():
    """Reads Windows System Metrics and returns screen heigth in pixel"""

    logging.debug('getScreenHeight()')

    height = win32api.GetSystemMetrics(1)
    logging.debug('getScreenHeight - height = {}'.format(height))
    return height

def addImageToDatabase(full_image_url, image_name, image_source):
    """Creates a database if it does not yet exist and writes full image url
    given by 'full_image_url' as primary key, image name given by 'image_name'
    and image source given by 'image_source' to a database
    """

    logging.debug('addImageToDatabase({}, {}, {})'.format(full_image_url, image_name, image_source))

    dir_path = os.path.join(os.environ['LOCALAPPDATA'],'WarietyWallpaperImages')
    os.makedirs(dir_path, exist_ok=True)
    db_file = os.path.join(dir_path,'wariety.db')
    conn = sqlite3.connect(db_file)
    c = conn.cursor()

    # Create table
    c.execute("""
        CREATE TABLE IF NOT EXISTS wallpapers (
        id integer primary key,
        iurl text unique,
        iname text,
        ipath text,
        isource text)
        """)

    # Insert a row of data
    c.execute("""INSERT INTO wallpapers (iurl, iname, isource)
        VALUES (?,?,?)""", (full_image_url, image_name, image_source))
    
    # Save (commit) the changes
    conn.commit()

    # Close connection
    conn.close()

def getImagePathFromDatabase(full_image_url):
    """Reads database and returns full image path based on full image url
    given by 'full_image_url'  
    """

    logging.debug('getImagePathFromDatabase({})'.format(full_image_url))

    dir_path = os.path.join(os.environ['LOCALAPPDATA'],'WarietyWallpaperImages')
    os.makedirs(dir_path, exist_ok=True)
    db_file = os.path.join(dir_path,'wariety.db')
    full_image_path = ""
    conn = sqlite3.connect(db_file)
    c = conn.cursor()

    # Create table
    c.execute("""
        CREATE TABLE IF NOT EXISTS wallpapers (
        id integer primary key,
        iurl text unique,
        iname text,
        ipath text,
        isource text)
        """)

    # Select a row
    c.execute("SELECT ipath FROM wallpapers WHERE iurl = ?", (full_image_url,))
    full_image_path = os.path.abspath(c.fetchone()[0])
    conn.close()
    logging.debug('getImagePathFromDatabase - full_image_path = {}'.format(full_image_path))
    return full_image_path

def updateImageInDatabase(full_image_url, full_image_path):
    """Updates image full path given by 'full_image_path' in database"""

    logging.debug('updateImageInDatabase({}, {})'.format(full_image_url, full_image_path))

    dir_path = os.path.join(os.environ['LOCALAPPDATA'],'WarietyWallpaperImages')
    os.makedirs(dir_path, exist_ok=True)
    db_file = os.path.join(dir_path,'wariety.db')
    conn = sqlite3.connect(db_file)
    c = conn.cursor()

    # Update a row
    c.execute("UPDATE wallpapers SET ipath = ? WHERE iurl = ?", (full_image_path, full_image_url))

    # Save (commit) the changes
    conn.commit()

    # Close connection
    conn.close()

def existsImageInDatabase(full_image_url):
    """Checks whether an image given by 'full_image_url' exists already in databse"""

    logging.debug('existsImageInDatabase({})'.format(full_image_url))

    dir_path = os.path.join(os.environ['LOCALAPPDATA'],'WarietyWallpaperImages')
    db_file = os.path.join(dir_path,'wariety.db')
    conn = sqlite3.connect(db_file)
    c = conn.cursor()

    # Create table
    c.execute("""
        CREATE TABLE IF NOT EXISTS wallpapers (
        id integer primary key,
        iurl text unique,
        iname text,
        ipath text,
        isource text)
        """)

    # Select a row
    c.execute("SELECT id FROM wallpapers WHERE iurl = ?", (full_image_url,))

    if c.fetchone() is not None:
        conn.close()
        logging.debug('existsImageInDatabase - True')
        return True
    else:
        conn.close()
        logging.debug('existsImageInDatabase - False')
        return False

def getGeneratedImageName(full_image_url):
    """Expects URL to an image, retrieves its file extension and returns
    an image name based on the current date and with the correct file
    extension
    """

    logging.debug('getGeneratedImageName({})'.format(full_image_url))

    image_name = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    image_extension = full_image_url.split(".")[-1]
    image_name = image_name + "." + image_extension
    logging.debug('getGeneratedImageName - image_name = {}'.format(image_name))
    return image_name

def downloadImage(full_image_url, image_name):
    """Creates the folder 'WarietyWallpaperImages' in the temporary
    locations if it does not yet exist. Downloads the image given
    by 'full_image_url', stores it there and returns the path to it
    """

    logging.debug('downloadImage({}, {})'.format(full_image_url, image_name))

    img_data = requests.get(full_image_url).content
    dir_path = os.path.join(os.environ['TEMP'],'WarietyWallpaperImages')
    os.makedirs(dir_path, exist_ok=True)
    with open(os.path.join(dir_path, image_name), 'wb') as handler:
        handler.write(img_data)
    logging.debug('downloadImage - dir_path = {}'.format(dir_path))
    logging.debug('downloadImage - image_name = {}'.format(image_name))
    return os.path.join(dir_path, image_name)

def getRandomImageFromDatabase():
    """Returns either full image path of a random image from the database or
    from any of the other image sources
    """

    logging.debug('getRandomImageFromDatabase()')

    dir_path = os.path.join(os.environ['LOCALAPPDATA'],'WarietyWallpaperImages')
    os.makedirs(dir_path, exist_ok=True)
    full_image_path = ""
    db_file = os.path.join(dir_path,'wariety.db')
    full_image_path = ""
    conn = sqlite3.connect(db_file)
    c = conn.cursor()

    # Create table
    c.execute("""
        CREATE TABLE IF NOT EXISTS wallpapers (
        id integer primary key,
        iurl text unique,
        iname text,
        ipath text,
        isource text)
        """)

    c.execute("SELECT id, ipath FROM wallpapers")

    result = c.fetchall()
    conn.close()

    max = len(result)
    try:
        choice = random.randint(0, int(max+max*.5))
        full_image_path = os.path.abspath(result[choice][1])
        logging.debug('getRandomImageFromDatabase - full_image_path = {}'.format(full_image_path))
        return full_image_path
    except:
        full_image_path = getRandomImageFromAnySource()
        logging.debug('getRandomImageFromDatabase - full_image_path = {}'.format(full_image_path))
        return full_image_path

def getRandomImageFromAnySource():
    """Returns full image path of any image source
    """

    logging.debug('getRandomImageFromAnySource()')

    myChoice = random.choice(['bing', 'flickr', 'spotlight', 'wikimedia'])
    if myChoice == 'bing':
        logging.debug('getRandomImageFromAnySource - getLatestBingWallpaperRemote()')
        return getLatestBingWallpaperRemote()
    elif myChoice == 'flickr':
        logging.debug('getRandomImageFromAnySource - getLatestFlickrWallpaperRemote()')
        return getLatestFlickrWallpaperRemote()
    elif myChoice == 'spotlight':
        logging.debug('getRandomImageFromAnySource - getLatestWallpaperLocal()')
        return getLatestWallpaperLocal()
    elif myChoice == 'wikimedia':
        logging.debug('getRandomImageFromAnySource - getLatestWikimediaWallpaperRemote()')
        return getLatestWikimediaWallpaperRemote()

def getLatestWikimediaWallpaperRemote():
    """Retrieves the URL of the latest image of Wikimedia Picture Of The Day,
    downloads the image, stores it in a temporary folder and returns the path
    to it
    """

    logging.debug('getLatestWikimediaWallpaperRemote()')

    # get image url
    response = requests.get("https://commons.wikimedia.org/wiki/Hauptseite")
    match = re.search('.*mainpage-potd.*src=\"([^\"]*)\".*', response.text)
    image_url = match.group(1)
    full_image_url = image_url.replace('500px','1920px')

    # image's name
    image_name = getGeneratedImageName(full_image_url)

    # Check and maintain DB
    if not existsImageInDatabase(full_image_url):
        addImageToDatabase(full_image_url, image_name, "wikimedia")
        # download and save image
        full_image_path = downloadImage(full_image_url, image_name)
        updateImageInDatabase(full_image_url, full_image_path)
    else:
        full_image_path = getImagePathFromDatabase(full_image_url)

    # Return full path to image
    logging.debug('getLatestWikimediaWallpaperRemote - full_image_path = {}'.format(full_image_path))
    return full_image_path

def getLatestFlickrWallpaperRemote():
    """Retrieves the URL of the latest image of Peter Levi's Flickr Collection,
    downloads the image, stores it in a temporary folder and returns the path
    to it
    """

    logging.debug('getLatestFlickrWallpaperRemote()')

    # get image url
    response = requests.get("https://www.flickr.com/photos/peter-levi/")
    match = re.search('([0-9]{10})_.*\.jpg\)', response.text)
    image_id = match.group(1)
    image_url = "https://www.flickr.com/photos/peter-levi/"+image_id+"/sizes/h/"
    response = requests.get(image_url)
    pattern = 'http.*'+image_id+'.*_h\.jpg'
    match = re.search(pattern, response.text)
    full_image_url = match.group(0)

    # image's name
    image_name = getGeneratedImageName(full_image_url)

    # Check and maintain DB
    if not existsImageInDatabase(full_image_url):
        addImageToDatabase(full_image_url, image_name, "flickr")
        # download and save image
        full_image_path = downloadImage(full_image_url, image_name)
        updateImageInDatabase(full_image_url, full_image_path)
    else:
        full_image_path = getImagePathFromDatabase(full_image_url)

    # Return full path to image
    logging.debug('getLatestFlickrWallpaperRemote - full_image_path = {}'.format(full_image_path))
    return full_image_path

def getLatestBingWallpaperRemote():
    """Retrieves the URL of Bing's Image Of The Day image, downloads the image,
    stores it in a temporary folder and returns the path to it
    """

    logging.debug('getLatestBingWallpaperRemote()')

    # get image url
    response = requests.get("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US")
    image_data = json.loads(response.text)

    image_url = image_data["images"][0]["url"]
    image_url = image_url.split("&")[0]
    full_image_url = "https://www.bing.com" + image_url
    logging.debug('getLatestBingWallpaperRemote - full_image_url = {}'.format(full_image_url))

    # image's name
    image_name = getGeneratedImageName(full_image_url)

    # Check and maintain DB
    if not existsImageInDatabase(full_image_url):
        addImageToDatabase(full_image_url, image_name, "bing")
        # download and save image
        full_image_path = downloadImage(full_image_url, image_name)
        updateImageInDatabase(full_image_url, full_image_path)
    else:
        full_image_path = getImagePathFromDatabase(full_image_url)

    # Return full path to image
    logging.debug('getLatestBingWallpaperRemote - full_image_path = {}'.format(full_image_path))
    return full_image_path

def usage(arg):
    """Shows help of this tool"""

    logging.debug('usage({})'.format(arg))

    myDocstring = ""
    if arg in ["-i", "--info"]:
        myDocstring = myDocstring+"\n"+__doc__
        myDocstring = myDocstring+"\n    AUTHOR\n\n        $author$\n"
        myDocstring = myDocstring+"\n    LICENSE\n\n        $license$\n"
        parser.print_help()
    elif arg in ["-v", "-i", "--version", "--info"]:
        myDocstring = myDocstring+"\n    VERSION\n\n        $version$\n"    
    myDocstring = myDocstring.replace('$version$', __version__)
    myDocstring = myDocstring.replace('$author$',__author__)
    myDocstring = myDocstring.replace('$license$',__license__)
    print(myDocstring)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Load and show nice Windows background images.')
    parser.add_argument('-b', '--bing', help = "set Bing Image Of The Day as wallpaper", action="store_true")
    parser.add_argument('-f', '--flickr', help = "set Peter Levi's Flickr Collection as wallpaper", action="store_true")
    parser.add_argument('-s', '--spotlight', help = "set Microsoft Spotlight as wallpaper [default]", action="store_true")
    parser.add_argument('-w', '--wikimedia', help = "set Wikimedia Picture Of The Day as wallpaper", action="store_true")
    parser.add_argument('-r', '--random', help = "set wallpaper from random source", action="store_true")
    parser.add_argument('-i','--info', help = "show license and author information", action="store_true")
    parser.add_argument('-v', '--version', help = "show version", action="store_true")
    parser.add_argument('-d','--debug', help = "write debug output to logfile", action="store_true")
    path = ""
    args = parser.parse_args()
    set_any_option = False
    if args.debug:
        myname = os.path.basename(__file__).split('.')[0]
        fname = os.path.join(os.path.dirname(os.path.realpath(__file__)), '{}.log'.format(myname))
        logging.basicConfig(filename=fname, filemode='a', level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        myargs = []
        for arg in vars(args):
            if getattr(args, arg):
                myargs.append('--{}'.format(arg))
        logging.debug('__main__ - Starting application with "{}.py {}"'.format(myname,' '.join(myargs)))
    if args.info:
        usage('-i')
        set_any_option = True
    if args.version:
        usage('-v')
        set_any_option = True
    if args.bing:
        path = getLatestBingWallpaperRemote()
        set_any_option = True
    if args.flickr:
        path = getLatestFlickrWallpaperRemote()
        set_any_option = True
    if args.spotlight:
        path = getLatestWallpaperLocal()
        set_any_option = True
    if args.random:
        path = getRandomImageFromDatabase()
        set_any_option = True
    if args.wikimedia:
        path = getLatestWikimediaWallpaperRemote()
        set_any_option = True
    if not set_any_option:
        # default
        path = getLatestWallpaperLocal()
    setWallpaperWithCtypes(path)
    logging.debug('__main__ - Stopping application with exit code "0"\n')
    sys.exit(0)