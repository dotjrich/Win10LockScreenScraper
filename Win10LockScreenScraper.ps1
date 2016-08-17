#
# Win10LockScreenScraper
# dotjrich
#

# As of 8/14/2016, this is where the images are stored.
$sourceDirectory = "$env:userprofile\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"

$destinationDirectory = "$env:userprofile\Desktop\LockScreenPics"
New-Item -ItemType Directory -Force -Path $destinationDirectory | Out-Null

Get-ChildItem $sourceDirectory |
ForEach-Object {
    $image = New-Object -ComObject Wia.ImageFile
    try {
        $image.loadfile($_.FullName)
    } catch [System.ArgumentException] {
        # Not all files here are images, so this will throw an error if such a file is encountered. We can safely ignore these.
    }

    $imageOrientation = ""

    # We only care about HD images... other images will be here (Windows Store icons, etc).
    # FYI: There are two versions of each image: 1) landscape, 2) portrait.
    if (($image.Width -eq 1920) -and ($image.Height -eq 1080)) {
        $imageOrientation = "landscape"
    } elseif (($image.Width -eq 1080) -and ($image.Height -eq 1920)) {
        $imageOrientation = "portrait"
    }

    # If we've found an image with the proper resolution and we haven't copied it, copy it.
    if ($imageOrientation -ne "") {
        # TODO: I've noticed that some of the images actually have information in the metadata (name, location, etc). We can use that for the file name.
        $destinationFile = "$destinationDirectory/$_-$imageOrientation.jpg"
        if (-not (Test-Path -Path $destinationFile)) {
            Copy-Item -Path $_.FullName -Destination $destinationFile
        }
    }
}