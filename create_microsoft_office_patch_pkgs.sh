#!/bin/bash

loggedInUser=$(/usr/bin/python -c 'from SystemConfiguration import SCDynamicStoreCopyConsoleUser; import sys; username = (SCDynamicStoreCopyConsoleUser(None, None, None) or [None])[0]; username = [username,""][username in [u"loginwindow", None, u""]]; sys.stdout.write(username + "\n");')
userHome=$(/usr/bin/dscl . read "/Users/$loggedInUser" NFSHomeDirectory | cut -c 19-)
filePath="$1"

function finish () {
    if [[ -n "$pkgName" ]] && [[ -d "/tmp/$pkgName" ]]; then
        rm -R "/tmp/$pkgName"
    fi
    exit "$1"
}

# Determine the MS Office PKG we are going to rebuild
[[ -z "$filePath" ]] && filePath=$(osascript -e 'tell app (path to frontmost application as Unicode text) to set new_file to POSIX path of (choose file with prompt "Choose Office|Word|Excel|PowerPoint|Outlook .pkg to rebuild." of type {"PKG"})' 2> /dev/null)
[[ -z "$filePath" ]] && echo "User cancelled; exiting." && finish 1
fileName=${filePath##*/}
pkgName=${fileName%.*}

# Expand the package at a temporary location
pkgutil --expand "$filePath" "/tmp/$pkgName"

# Read informatoin about the PKG
if [[ -f "/tmp/$pkgName/Distribution" ]]; then
    pkgInfo=$(grep pkg-ref "/tmp/$pkgName/Distribution" | grep id | grep version | grep -v autoupdate | grep -v licensing | head -1)
    # Get the ID of the package
    id=$(echo "$pkgInfo" | tr " " "\n" | grep id | grep -o '".*"' | cut -d \" -f2)

    # Get the version of the package
    version=$(echo "$pkgInfo" | tr " " "\n" | grep version | grep -o '".*"' | cut -d \" -f2)
    majorVersion=$(awk -F '.' '{print $1}' <<< "$version")
    minorVersion=$(awk -F '.' '{print $2}' <<< "$version")


    # Make sure we obtain the specifics of the package
    [[ -z "$id" ]] && echo "Error; ID of package is null; exiting." && finish 1
    [[ -z "$version" ]] && echo "Error; Version of package is null; exiting." && finish 1
    echo "Package ID: $id;"
    echo "Package version: $version; Major: $majorVersion; Minor: $minorVersion;"
else
    echo "Error, Distrbution file does not exist; exiting."
    finish 1
fi

if [[ ! "$id" =~ office|powerpoint|outlook|excel|word ]]; then
    echo "$id is not a supported MS Office product for this process; exiting."
    finish 1
fi

# Create the PKG source files
rootDir="$userHome/${id}_${version}"
mkdir -p "$rootDir/scripts"
mkdir -p "$rootDir/build"

# Determine the VL Serializer version to use
if [[ "$majorVersion" == "16" ]] && [[ "$minorVersion" -le "16" ]] || [[ "$majorVersion" == "15" ]]; then
    serializer="Microsoft_Office_2016_VL_Serializer_2.0.pkg"
    officeVersion="2016"
elif [[ "$majorVersion" == "16" ]] && [[ "$minorVersion" -ge "17" ]]; then
    serializer="Microsoft_Office_2019_VL_Serializer.pkg"
    officeVersion="2019"
else
    echo "Could not determine VL Serializer to use; exiting."
    finish 1
fi

# Save the VL Serializer for later use and copy it to the scripts directory for the PKG
appSupportDir="$userHome/Library/Application Support/Create_Microsoft_Update_PkGs"
if [[ -f "$appSupportDir/$serializer" ]]; then
    echo "Found serializer in $appSupportDir; using that for new package..."
    cp "$appSupportDir/$serializer" "$rootDir/scripts/"
else
    findSerializer=$(osascript -e "tell app (path to frontmost application as Unicode text) to set new_file to POSIX path of (choose file with prompt \"Browse to your $serializer package.\" of type {\"PKG\"})" 2> /dev/null)
    [[ -z "$findSerializer" ]] && echo "User cancelled; exiting." && finish 1
    mkdir -p "$appSupportDir"
    echo "Saving the Serializer to $appSupportDir for later use..."
    cp "$findSerializer" "$appSupportDir/$serializer"
    cp "$findSerializer" "$rootDir/scripts/$serializer"
fi

# Create the postinstall script for the PKG
cat << EOF > "$rootDir/scripts/postinstall"
#!/bin/bash

# Determine working directory
install_dir=\$(dirname "\$0")

# Install the MS Office product
/usr/sbin/installer -dumplog -verbose -pkg "\$install_dir/${pkgName}.pkg" -target "\$3"

# Install the VL_Serializer
/usr/sbin/installer -dumplog -verbose -pkg "\$install_dir/$serializer" -target "\$3"

exit 0
EOF

# Move the Office PKG to the scripts directory
cp "${filePath}" "$rootDir/scripts/"
chmod +x "$rootDir/scripts/postinstall"

# Build the PKG
/usr/bin/pkgbuild --nopayload --scripts "/$rootDir/scripts" --identifier "$id" --version "$version" "$rootDir/build/${pkgName}_${officeVersion}_Repackaged.pkg"

if [[ -f "$rootDir/build/${pkgName}_${officeVersion}_Repackaged.pkg" ]]; then
    echo "Successfully created package."
    open -R "$rootDir/build/${pkgName}_${officeVersion}_Repackaged.pkg"
else
    echo "Error, package failed to build; exiting."
    finish 1
fi

finish 0
