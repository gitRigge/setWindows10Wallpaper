# setWindows10Wallpaper
Sets Windows 10 wallpaper to Microsoft Spotlight or Bing Image Of The Day

## DESCRIPTION

This tool can either set the latest locally stored Microsoft
Spotlight (--spotlight) image as Desktop wallpaper. Or it fetches
the latest image from Bing's Image Of The Day (--bing) collection
and set's it as wallpaper. Or it does randomly one of the two (--random)
Default: Microsoft Spotlight

## USAGE

    usage: setWindows10Wallpaper_cli.py [-h] [-b] [-f] [-s] [-r] [-w] [-i] [-v]

    Sets Windows 10 wallpaper to Microsoft Spotlight or Bing Image Of The Day.

    optional arguments:
      -h, --help       show this help message and exit
      -b, --bing       set Bing Image Of The Day as wallpaper
      -f, --flickr     set Peter Levi's Flickr Collection as wallpaper
      -s, --spotlight  set Microsoft Spotlight as wallpaper [default]
      -w, --wikimedia  set Wikimedia Picture Of The Day as wallpaper
      -r, --random     set wallpaper from random source
      -i, --info       show license and author information
      -v, --version    show version
      -d, --debug      write debug output to logfile