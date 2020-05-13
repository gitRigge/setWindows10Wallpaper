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
__version__ = "1.0"
__status__ = "Development"

proxies = {
  'http': 'http://0.0.0.0:80/',
  "https": "http://0.0.0.0:8080",
}

def set_wallpaper_with_ctypes(path):
    """Sets asset given by 'path' as current Desktop wallpaper"""

    logging.debug('setWallpaperWithCtypes({})'.format(path))

    cs = ctypes.create_string_buffer(path.encode('utf-8'))
    ok = ctypes.windll.user32.SystemParametersInfoA(win32con.SPI_SETDESKWALLPAPER, 0, cs, 0)

def get_latest_wallpaper_local():
    """Loops through all locally stored Windows Spotlight assets
    and copies and returns the latest asset which has the same orientation as the screen
    """

    logging.debug('get_latest_wallpaper_local()')

    list_of_files = sorted(glob.glob(os.environ['LOCALAPPDATA']
        +r'\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets'
        +'\*'), key=os.path.getmtime, reverse=True)
    for asset in list_of_files:
        extension = imghdr.what(asset)
        if extension in ['jpeg', 'jpg', 'png']:
            if is_image_landscape(asset) == is_screen_landscape():
                # Generate pseudo url
                full_image_url = os.path.split(asset)[1]
                if not exists_image_in_database(full_image_url):
                    image_name = get_generated_image_name(asset+'.'+extension)
                    add_image_to_database(full_image_url, image_name, "spotlight")
                    dir_path = os.path.join(os.environ['TEMP'],'WarietyWallpaperImages')
                    os.makedirs(dir_path, exist_ok=True)
                    full_image_path = os.path.join(dir_path, image_name)
                    shutil.copyfile(asset, full_image_path)
                    update_image_in_database(full_image_url, full_image_path)
                    logging.debug('get_latest_wallpaper_local - full_image_path = {}'.format(full_image_path))
                    return full_image_path
                else:
                    logging.debug('get_latest_wallpaper_local - get_tmage_path_from_database({})'.format(full_image_url))
                    return get_image_path_from_database(full_image_url)

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

def is_image_landscape(asset):
    """Checks the orientation of the asset given by 'asset' and returns 'True' if the asset's
    orientation is landscape
    """
    
    logging.debug('is_image_landscape({})'.format(asset))

    myDim = get_image_size(asset)
    # Calculate Width:Height; > 0 == landscape; < 0 == portrait
    if myDim[0]/myDim[1] > 1:
        logging.debug('is_image_landscape - True')
        return True
    else:
        logging.debug('is_image_landscape - False')
        return False

def is_screen_landscape():
    """Checks the current screen orientation and returns 'True' if the screen orientation
    is landscape
    """

    logging.debug('is_screen_landscape()')

    if get_screen_width()/get_screen_height() > 1:
        logging.debug('is_image_landscape - True')
        return True
    else:
        logging.debug('is_image_landscape - False')
        return False

def get_screen_width():
    """Reads Windows System Metrics and returns screen width in pixel"""

    logging.debug('get_screen_width()')

    width = win32api.GetSystemMetrics(0)
    logging.debug('get_screen_width - width = {}'.format(width))
    return width

def get_screen_height():
    """Reads Windows System Metrics and returns screen heigth in pixel"""

    logging.debug('get_screen_height()')

    height = win32api.GetSystemMetrics(1)
    logging.debug('get_screen_height - height = {}'.format(height))
    return height

def add_image_to_database(full_image_url, image_name, image_source):
    """Creates a database if it does not yet exist and writes full image url
    given by 'full_image_url' as primary key, image name given by 'image_name'
    and image source given by 'image_source' to a database
    """

    logging.debug('add_image_to_database({}, {}, {})'.format(full_image_url, image_name, image_source))

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

def database_maintenance():
    """Keep database and image folder synced"""

    logging.debug('database_maintenance()')

    all_imagepaths = get_all_images_from_database()
    for imagepath in all_imagepaths:
        if not os.path.isfile(imagepath):
            delete_image_from_database(imagepath)

def get_all_images_from_database():
    """Reads database and returns full image path of all
    images currently stored in the database
    """

    logging.debug('get_all_images_from_database()')

    dir_path = os.path.join(os.environ['LOCALAPPDATA'],'WarietyWallpaperImages')
    os.makedirs(dir_path, exist_ok=True)
    db_file = os.path.join(dir_path,'wariety.db')
    full_image_paths = []
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
    c.execute("SELECT ipath FROM wallpapers", ())
    result = c.fetchall()
    conn.close()
    for item in result:
        full_image_paths.append(os.path.abspath(item[0]))
    logging.debug('get_all_images_from_database - full_image_paths = {}'.format(full_image_paths))
    return full_image_paths

def delete_image_from_database(full_image_path):
    """Deletes image given by 'full_image_path' from database
    """

    logging.debug('delete_image_from_database({})'.format(full_image_path))

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

    # Select a row
    c.execute("DELETE FROM wallpapers WHERE ipath = ?", (full_image_path,))
    conn.commit()
    conn.close()

def get_image_path_from_database(full_image_url):
    """Reads database and returns full image path based on full image url
    given by 'full_image_url'
    """

    logging.debug('get_image_path_from_database({})'.format(full_image_url))

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
    logging.debug('get_image_path_from_database - full_image_path = {}'.format(full_image_path))
    return full_image_path

def update_image_in_database(full_image_url, full_image_path):
    """Updates image full path given by 'full_image_path' in database"""

    logging.debug('update_image_in_database({}, {})'.format(full_image_url, full_image_path))

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

def exists_image_in_database(full_image_url):
    """Checks whether an image given by 'full_image_url' exists already in databse"""

    logging.debug('exists_image_in_database({})'.format(full_image_url))

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
        logging.debug('exists_image_in_database - True')
        return True
    else:
        conn.close()
        logging.debug('exists_image_in_database - False')
        return False

def get_generated_image_name(full_image_url):
    """Expects URL to an image, retrieves its file extension and returns
    an image name based on the current date and with the correct file
    extension
    """

    logging.debug('get_generated_image_name({})'.format(full_image_url))

    image_name = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    image_extension = full_image_url.split(".")[-1]
    image_name = image_name + "." + image_extension
    logging.debug('get_generated_image_name - image_name = {}'.format(image_name))
    return image_name

def download_image(full_image_url, image_name):
    """Creates the folder 'WarietyWallpaperImages' in the temporary
    locations if it does not yet exist. Downloads the image given
    by 'full_image_url', stores it there and returns the path to it
    """

    logging.debug('download_image({}, {})'.format(full_image_url, image_name))

    if use_proxy:
        img_data = requests.get(full_image_url, proxies=proxies, timeout=5, verify=False)
    else:
        img_data = requests.get(full_image_url).content
    dir_path = os.path.join(os.environ['TEMP'],'WarietyWallpaperImages')
    os.makedirs(dir_path, exist_ok=True)
    with open(os.path.join(dir_path, image_name), 'wb') as handler:
        handler.write(img_data)
    logging.debug('download_image - dir_path = {}'.format(dir_path))
    logging.debug('download_image - image_name = {}'.format(image_name))
    return os.path.join(dir_path, image_name)

def get_random_image_from_database():
    """Returns either full image path of a random image from the database or
    from any of the other image sources
    """

    logging.debug('get_random_image_from_database()')

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
        logging.debug('get_random_image_from_database - full_image_path = {}'.format(full_image_path))
        return full_image_path
    except:
        full_image_path = get_random_image_from_any_source()
        logging.debug('get_random_image_from_database - full_image_path = {}'.format(full_image_path))
        return full_image_path

def get_random_image_from_any_source():
    """Returns full image path of any image source
    """

    logging.debug('get_random_image_from_any_source()')

    myChoice = random.choice(['bing', 'bingarchive', 'flickr', 'spotlight', 'wikimedia'])
    if myChoice == 'bing':
        logging.debug('get_random_image_from_any_source - get_latest_bing_wallpaper_remote()')
        return get_latest_bing_wallpaper_remote()
    elif myChoice == 'bingarchive':
        logging.debug('get_random_image_from_any_source - get_a_bing_archive_wallpaper_remote()')
        return get_a_bing_archive_wallpaper_remote()
    elif myChoice == 'flickr':
        logging.debug('get_random_image_from_any_source - get_latest_flickr_wallpaper_remote()')
        return get_latest_flickr_wallpaper_remote()
    elif myChoice == 'spotlight':
        logging.debug('get_random_image_from_any_source - getLatestWallpaperLocal()')
        return get_latest_wallpaper_local()
    elif myChoice == 'wikimedia':
        logging.debug('get_random_image_from_any_source - get_latest_wikimedia_wallpaper_remote()')
        return get_latest_wikimedia_wallpaper_remote()

def get_a_bing_archive_wallpaper_remote():
    """Retrieves the URL of one image of Bing Wallpaper Archive,
    downloads the image, stores it in a temporary folder and returns the path
    to it
    """
    
    logging.debug('get_a_bing_archive_wallpaper_remote()')

    now = datetime.datetime.now()
    url = "https://bingwallpaper.anerg.com/de/{}".format(now.strftime('%Y%m'))

    # get image url
    if use_proxy:
        
        response = requests.get(url, proxies=proxies, timeout=15, verify=False)
    else:
        response = requests.get(url)
    match = re.findall('.*src=\"([^\"]*\.jpg)\".*', response.text)
    for i in range(0, len(match)):
        full_image_url = "https:{}".format(match[i])
        
        # image's name
        image_name = get_generated_image_name(full_image_url)
        
        # Check and maintain DB
        if not exists_image_in_database(full_image_url) and i+1 < len(match):
            add_image_to_database(full_image_url, image_name, "bingarchive")
            # download and save image
            full_image_path = download_image(full_image_url, image_name)
            update_image_in_database(full_image_url, full_image_path)

            # Return full path to image
            logging.debug('get_a_bing_archive_wallpaper_remote - full_image_path = {}'.format(full_image_path))
            return full_image_path
        elif i+1 == len(match):
            full_image_path = get_image_path_from_database(full_image_url)

            # Return full path to image
            logging.debug('get_a_bing_archive_wallpaper_remote - full_image_path = {}'.format(full_image_path))
            return full_image_path

def get_latest_wikimedia_wallpaper_remote():
    """Retrieves the URL of the latest image of Wikimedia Picture Of The Day,
    downloads the image, stores it in a temporary folder and returns the path
    to it
    """

    logging.debug('get_latest_wikimedia_wallpaper_remote()')

    # get image url
    if use_proxy:
        response = requests.get("https://commons.wikimedia.org/wiki/Hauptseite", proxies=proxies, timeout=15, verify=False)
    else:
        response = requests.get("https://commons.wikimedia.org/wiki/Hauptseite")
    match = re.search('.*mainpage-potd.*src=\"([^\"]*)\".*', response.text)
    image_url = match.group(1)
    full_image_url = image_url.replace('500px','1920px')

    # image's name
    image_name = get_generated_image_name(full_image_url)

    # Check and maintain DB
    if not exists_image_in_database(full_image_url):
        add_image_to_database(full_image_url, image_name, "wikimedia")
        # download and save image
        full_image_path = download_image(full_image_url, image_name)
        update_image_in_database(full_image_url, full_image_path)
    else:
        full_image_path = get_image_path_from_database(full_image_url)

    # Return full path to image
    logging.debug('get_latest_wikimedia_wallpaper_remote - full_image_path = {}'.format(full_image_path))
    return full_image_path

def get_latest_flickr_wallpaper_remote():
    """Retrieves the URL of the latest image of Peter Levi's Flickr Collection,
    downloads the image, stores it in a temporary folder and returns the path
    to it
    """

    logging.debug('get_latest_flickr_wallpaper_remote()')

    # get image url
    if use_proxy:
        response = requests.get("https://www.flickr.com/photos/peter-levi/", proxies=proxies, timeout=5, verify=False)
    else:
        response = requests.get("https://www.flickr.com/photos/peter-levi/")
    match = re.search('([0-9]{10})_.*\.jpg\)', response.text)
    image_id = match.group(1)
    image_url = "https://www.flickr.com/photos/peter-levi/"+image_id+"/sizes/h/"
    if use_proxy:
        response = requests.get(image_url, proxies=proxies, timeout=5, verify=False)
    else:
        response = requests.get(image_url)
    pattern = 'http.*'+image_id+'.*_h\.jpg'
    match = re.search(pattern, response.text)
    full_image_url = match.group(0)

    # image's name
    image_name = get_generated_image_name(full_image_url)

    # Check and maintain DB
    if not exists_image_in_database(full_image_url):
        add_image_to_database(full_image_url, image_name, "flickr")
        # download and save image
        full_image_path = download_image(full_image_url, image_name)
        update_image_in_database(full_image_url, full_image_path)
    else:
        full_image_path = get_image_path_from_database(full_image_url)

    # Return full path to image
    logging.debug('get_latest_flickr_wallpaper_remote - full_image_path = {}'.format(full_image_path))
    return full_image_path

def get_latest_bing_wallpaper_remote():
    """Retrieves the URL of Bing's Image Of The Day image, downloads the image,
    stores it in a temporary folder and returns the path to it
    """

    logging.debug('get_latest_bing_wallpaper_remote()')

    # get image url
    if use_proxy:
        response = requests.get("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US", proxies=proxies, timeout=5, verify=False)
    else:
        response = requests.get("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US")
    image_data = json.loads(response.text)

    image_url = image_data["images"][0]["url"]
    image_url = image_url.split("&")[0]
    full_image_url = "https://www.bing.com" + image_url
    logging.debug('get_latest_bing_wallpaper_remote - full_image_url = {}'.format(full_image_url))

    # image's name
    image_name = get_generated_image_name(full_image_url)

    # Check and maintain DB
    if not exists_image_in_database(full_image_url):
        add_image_to_database(full_image_url, image_name, "bing")
        # download and save image
        full_image_path = download_image(full_image_url, image_name)
        update_image_in_database(full_image_url, full_image_path)
    else:
        full_image_path = get_image_path_from_database(full_image_url)

    # Return full path to image
    logging.debug('get_latest_bing_wallpaper_remote - full_image_path = {}'.format(full_image_path))
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
    parser.add_argument('-a', '--bingarchive', help = "set Bing Wallpaper Archive as wallpaper", action="store_true")
    parser.add_argument('-f', '--flickr', help = "set Peter Levi's Flickr Collection as wallpaper", action="store_true")
    parser.add_argument('-s', '--spotlight', help = "set Microsoft Spotlight as wallpaper [default]", action="store_true")
    parser.add_argument('-w', '--wikimedia', help = "set Wikimedia Picture Of The Day as wallpaper", action="store_true")
    parser.add_argument('-r', '--random', help = "set wallpaper from random source", action="store_true")
    parser.add_argument('-p','--proxy', help = "use proxy to grab images", action="store_true")
    parser.add_argument('-i','--info', help = "show license and author information", action="store_true")
    parser.add_argument('-v', '--version', help = "show version", action="store_true")
    parser.add_argument('-d','--debug', help = "write debug output to logfile", action="store_true")
    path = ""
    args = parser.parse_args()
    use_proxy = False
    set_any_option = False
    if args.debug:
        myname = os.path.basename(__file__).split('.')[0]
        mypath = os.path.join(os.environ['LOCALAPPDATA'],'WarietyWallpaperImages')
        fname = os.path.join(mypath, '{}.log'.format(myname))
        logging.basicConfig(filename=fname, filemode='a', level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        myargs = []
        for arg in vars(args):
            if getattr(args, arg):
                myargs.append('--{}'.format(arg))
        logging.debug('__main__ - Starting application with "{}.py {}"'.format(myname,' '.join(myargs)))
    if args.proxy:
        use_proxy = True
    # do maintenance in any case; do it only after debug
    database_maintenance()
    if args.info:
        usage('-i')
        set_any_option = True
    if args.version:
        usage('-v')
        set_any_option = True
    if args.bing:
        path = get_latest_bing_wallpaper_remote()
        set_any_option = True
    if args.bingarchive:
        path = get_a_bing_archive_wallpaper_remote()
        set_any_option = True
    if args.flickr:
        path = get_latest_flickr_wallpaper_remote()
        set_any_option = True
    if args.spotlight:
        path = get_latest_wallpaper_local()
        set_any_option = True
    if args.random:
        path = get_random_image_from_database()
        set_any_option = True
    if args.wikimedia:
        path = get_latest_wikimedia_wallpaper_remote()
        set_any_option = True
    if not set_any_option:
        # default
        path = get_latest_wallpaper_local()
    set_wallpaper_with_ctypes(path)
    logging.debug('__main__ - Stopping application with exit code "0"\n')
    sys.exit(0)