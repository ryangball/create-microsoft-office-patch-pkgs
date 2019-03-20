# Create Microsoft Office Patch PKGs
Easily create individual Microsoft Office component packages to be used with Jamf Pro’s Patch Management.

### Explanation
Patch Management within Jamf Pro requires individual packages when patching the component applications within the suite (PowerPoint, Outlook, Excel, Word). Creating these individual packages can be tedious especially if you have Volume Licensing for Office.

This script will either take an Office component installer/updater of your choosing, and automatically re-package it with your VL_Serializer.pkg included.

### Features
- Allows you to specify a path as a parameter or displays a dialog to select an Office component installer/updater
- Determines the version of the component installer/updater
- Determines which VL_Serializer.pkg should be included (2016/2019) and prompts you to select it
- Saves the VL_Serializer.pkg to a specific location on the system to eliminate future prompts to select it
- Appends the name of the PKG with "_Repackaged.pkg"
- Reveals the repackaged PKG in Finder when done

### Run With Path as Parameter
You can run this script with a parameter if you choose like so:
```
./create_microsoft_office_patch_pkgs.sh "~/Desktop/Microsoft_Word_16.23.19030902_Installer.pkg"
```
