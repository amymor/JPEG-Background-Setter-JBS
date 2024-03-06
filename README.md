# JPEG Background Setter(JPS)
Set your Actual JPEG image as background.

## Why JBS?
When you set a JPEG image as your desktop background, Windows converts it to a `TranscodedWallpaper`, reducing the quality by 85% by default. Then, it converts it again to a JPEG image, reducing the quality by an additional 90%. This results in a final quality of 76.5% of the original. We refer to this entire process as `Double Conversion`, and the second conversion as `Horrible Second Conversion`.
The quality reduction percentage of the first conversion is controlled by `JPEGImportQuality` in the registry, but we have no control over the second conversion in the registry. Even if we control both conversions and set both to 100%, we can't avoid quality reduction because the JPEG format is a `lossy` format, so Windows itself is trying the wrong way by converting JPEG images and Microsoft should give users an option to avoid this. Fortunately, with JBS, we can get rid of both conversions.

### Comparison
Open the `Comparison\index.html` in your browser, and check the `Comparison\guide.txt` for more information.


### Cons
1. For slideshows, currently, there is no option to automatically change the background after a period of time like Windows slideshow. However, you can do it manually by using the desktop context menu. (If I have free time and receive positive reactions from users, I will implement auto-change too)

### Pros
1. For slideshows, you have the "Previous desktop background" in your desktop context menu. (To-Do)
2. The `Set as Desktop background` in the JPEG context menu will change to set your actual JPEG image as the background without any conversion. (To-Do)
3. Preserves the exact quality of your JPEG images, just like when you set a PNG image as the background.
4. Reduces read/write to the system disk, which is beneficial for SSD users with a lower time span. (Avoids the creation of unnecessary cache files and also avoids unnecessary increases in the size of `TranscodedWallpaper` when the quality of your JPEG image is less than the value of `JPEGImportQuality`)
5. The value of JPEGImportQuality doesn't matter anymore. (Note: `JPEGImportQuality` is not present in the registry by default)
6. Full support for `Webp` image format unlike Windows Slideshow.

## Download the latest
Initial release:

## JBS Usage
```
JBS.exe /? (To show this help)

JBS.exe "Wallpaper_Path"
JBS.exe /u (To Uninstall)

Wallpaper position setting would read from JBS.ini:
Position_Setting_Number = N

The Position_Setting_Number should be the following numbers:
0 - Center
1 - Tile
2 - Stretch
6 - Fit
10 - Fill
22 - Span
```

## Set a picture as background without quality reduction
Drag and drop the image file on `JBS.exe` or pass image path as first parameter like:
```
JBS.exe "C:\Images\MyDesktopBackground.jpg"
```
Note: make sure to use double quote (`"`) in start and end of image path

### To-Do - Replace Set as Desktop Background Context Menu
1. Run the `2.Install-Context-Menu.bat` and select 1.
2. Right-click on any image, you will see a new option (`Set as Background with JBS`) which replaces the old one.


## Set folder pictures as background without quality reduction
1. Drag and drop the folder that contains pictures on `1.Slideshow.bat`, or create a `FolderPath.ini` that contains your pictures folder path in first line, example:
```
C:\Path\to\ImagesFolder
```

3. The script will ask you which picture to set as the background image by picture number alphabetically, so if there are 10 picture in your folder, entering the number 2 will set the second picture as the wallpaper.

Note: If `%1` (passed parameter like drag and drop) and `FolderPath.ini"` do not exist, the script will ask you for a custom folder path. It will then create a `FolderPath.ini` containing that path. Therefore, you don't need to specify the folder address again in subsequent runs.

## JBS GUI (Recommended)
### `JBSGUI.exe` is an application that functions similarly to `1.Slideshow.bat`, offering more control over the current background.

### To-Do Desktop Context Menu
1. Run the `2.Install-Context-Menu.bat` and select 2.
2. Right-click on your desktop, and you will see two options:
`Next Desktop Background`
`Previous Desktop Background`

## Notes
1. If you change your wallpaper using `Windows Personalization` settings, whether it's through UWP Personalization or the old Classic Personalization, the changes will most likely revert to the default.


# Uninstallation
1. Run the `Uninstall.bat`.
2. Set any picture as your desktop background as usual.
