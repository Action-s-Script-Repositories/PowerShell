#!/bin/bash
#Convert-Package.sh

helpFunction()
{
    echo ""
    echo "Usage: $0 -a parameterA -b parameterB -c parameterC"
    echo -e "\t-a Description of what is parameterA"
    echo -e "\t-b Description of what is parameterB"
    echo -e "\t-c Description of what is parameterC"
    exit 1 # Exit script after printing help
}

while getopts "a:b:c:" opt
do
    case "$opt" in
        a ) parameterA="$OPTARG" ;;
        b ) parameterB="$OPTARG" ;;
        c ) parameterC="$OPTARG" ;;
        ? ) helpFunction ;; # Print helpFunction in case parameter is non-existent
    esac
done

# Print helpFunction in case parameters are empty
if [ -z "$paramterA" ] || [ -z "$parameterB" ] || [ -z "$parameterC" ]
then
    echo "Some or all of the parameters are empty";
    helpFunction
fi

# Begin script in case all parameters are set

### Mount dmg
# hdiutil attach [path to diskimage.dmg]
    ### Install app
    ### Unmount dmg
    # hdiutil detach /dev/[disk]

### Convert app to pkh
# sudo productbuild --component [path to installed app] [path to converted pkg]
## sudo productbuild --component ../../Applications/VMware\ Horizon\ Client.app ./Downloads/PackageApps/Horizon\ View/build/HorizonViewConverted.pkg

### Sign pkg
# productsign --sign "Developer ID Installer: Jelle van de Kamp (3JHWZ2S9AC)" ./PackageApps/Horizon\ View/build/HorizonViewConverted.pkg ./PackageApps/Horizon\ View/build/HorizonViewConvertedSigned.pkg

### Convert pkg to intunemax
# ./intune-app-wrapping-tool-mac-1.2/IntuneAppUtil -c ./PackageApps/Hotizon\ View/build/HorizonViewConvertedSigned.pkg -o ./PackageApps/Horizon\ View/build


