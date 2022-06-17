# CreateManifestWithAppArguments.ps1
#
# Description:
# Creates modified AppxManifest.xml (AppxManifestNew.xml) to facilitate passing of command line arguments 
# from desktop applications coverted using the MSIX Packaging Tool
#
# Extracts command line arguements from application shortcuts (.lnk). 
# Adds the 'Parameter' attribute to the 'Application' node and set to the corresponding application arguments
# Adding additional schema definition namespace to support the Parameter attribute
#
# Requires:
# AppxManifest.xml and Shortcuts from packaged app. Shortcuts can be anywhere in the package.
#
# Output:
# AppxManifestNew.xml 
# Contents of AppxManifestNew.xml  will be added to edited MSIX package.
# Does not require Package Support Framework!

$application = @{
    id         = ''
    executable = ''
    arguments  = ''
    description = ''
}
$configjsonArray = [System.Collections.ArrayList]@()

Function StripQuotes {
    Param (
        [String]$StringWithQuotes
    )
    $strRes = $StringWithQuotes;
    if ($StringWithQuotes.IndexOf('"') -eq 0) {
        $p = $StringWithQuotes.Split('"')
        if ($p[1].Length -gt 1) {
            $strRes = $p[1];
        }
    }
    return $strRes;
}
Function GetCorrectPackagePath {
    Param (
        [String]$FileNameWithPath
    )
    $newPath = $FileNameWithPath;
    $fileFullPath = Get-ChildItem -Recurse (Split-Path $FileNameWithPath -Leaf) | Select-Object -ExpandProperty FullName
    if ($fileFullPath.Length -eq 0) {
        return $FileNameWithPath;
    }
    $vfsFilePath= $($fileFullPath -Split "VFS")[1]
   
    return '.' + $vfsFilePath;
    
}

    Function ParseShortCuts {
    Param (
        [System.Collections.ArrayList]$configjsonArray
    )
    
    # For each shortcut (lnk):
    # Extract the executable shortcut and detemine the index in the Applications collection in the manifest
    # https://www.alexandrumarin.com/add-shortcut-arguments-in-msix-with-psf/
    $sh = New-Object -ComObject WScript.Shell
    $files = Get-ChildItem -Recurse *.lnk
    foreach ($file in $files) {
        # if shortcut has no arguments, skip
        if ($sh.CreateShortcut($file).Arguments -eq "") {
            continue;
        }
        $object = new-object psobject -Property $application
        $object.id = [System.IO.Path]::GetFileNameWithoutExtension( $sh.CreateShortcut($file).TargetPath)
        $object.executable = $(Split-Path $sh.CreateShortcut($file).TargetPath -Leaf).ToLower()
        $oArgs = StripQuotes($sh.CreateShortcut($file).Arguments);
        $oArgs = GetCorrectPackagePath ($oArgs);
        $object.arguments = $oArgs;
        $object.description = $sh.CreateShortcut($file).description;
        $configjsonArray.Add($object);
    }
}

## Start

# <Application Id="VLC" Executable="VFS\ProgramFilesX64\VideoLAN\VLC\vlc.exe" uap10:Parameters="--no-qt-privacy-ask" EntryPoint="Windows.FullTrustApplication">
# https://docs.microsoft.com/en-us/uwp/schemas/appxpackage/uapmanifestschema/element-application

[xml]$manifest = get-content "AppxManifest.xml"
$appsInManifest= $manifest.Package.Applications.Application

# Add namespaces needed by Paramters
$nsm = New-Object System.Xml.XmlNamespaceManager($manifest.nametable)
$nsList = @( ("uap10", "http://schemas.microsoft.com/appx/manifest/uap/windows10/10") , ("desktop", "http://schemas.microsoft.com/appx/manifest/desktop/windows10") )
$ignoreNS = $manifest.Package.IgnorableNamespaces;

foreach ($ns in $nsList
) {
    $nsm.AddNamespace($ns[0], $ns[1])
    $manifest.Package.SetAttribute("xmlns:" + $ns[0], $ns[1])
    $ignoreNS = $ignoreNS + " " + $ns[0]
}
$manifest.Package.RemoveAttribute("IgnorableNamespaces")
$manifest.Package.SetAttribute("IgnorableNamespaces", $ignoreNS)

ParseShortCuts($configjsonArray)

foreach ($app in $appsInManifest) {
    if ($app.VisualElements.AppListEntry -eq "none") {
        continue
    }
    
    foreach ($shortcutApp in $configjsonArray) {
        if ($shortcutApp.executable -eq (Split-Path $app.Executable -Leaf)) {
            $newApp = $app.Clone();
            $newApp.Id = $app.Id + "1";
            $newApp.SetAttribute("Parameters", $nsm.LookupNamespace("uap10"), $shortcutApp.arguments);
            $manifest.Package.Applications.AppendChild($newApp);    
        }
    }
}


$appListManifest = [System.Collections.ArrayList]@()
foreach ($app in $appsInManifest) {
    $appListManifest.Add($app.Executable);
}

foreach ($app in $configjsonArray) {

    $searchResults = ($appListManifest | Where-Object { (Split-Path $app.executable -Leaf) -eq (Split-Path $_ -Leaf) })
    if ($searchResults.Length -eq 0) {
        #$newApp = $manifest.CreateElement("Application");
        $newApp = $manifest.Package.Applications.FirstChild.Clone();
        $newApp.SetAttribute("Executable",$app.executable);
        $newApp.SetAttribute("Id",$(($app.id + "1").ToUpper()));
        $newApp.SetAttribute("Parameters", $nsm.LookupNamespace("uap10"), $app.arguments);
        $newApp.SetAttribute("EntryPoint","Windows.FullTrustApplication");
        $newApp.VisualElements.SetAttribute("DisplayName", $app.description);
        $newApp.VisualElements.SetAttribute("Description", $app.description);
        $manifest.Package.Applications.AppendChild($newApp);    
    } 
}




# Write new manifest file
$manifest.Save($pwd.Path + "\AppxManifestNew.xml");




# $index = 0;
# Foreach ($app in $configjsonArray) {
#     $appPath = ($pwd.Path + "\" + $app.Executable)
#     try {
#         ExtractIcon -exePath ( $appPath ) -destinationFolder ( $desinationIconPath ) -iconName ( $appsInManifest.VisualElements.DisplayName )
#     } catch {
#         Write-Host $Error[0]
#     }
#     if (Test-Path($($desinationIconPath + "\" + $appsInManifest.VisualElements.DisplayName + ".ico")) ) {
#         Write-Host $("Icon successfully created: " + $desinationIconPath + "\" + $appsInManifest.VisualElements.DisplayName + ".ico")
#     } else {
#         Write-Host $("Icon NOT created: " + $desinationIconPath + "\" + $appsInManifest.VisualElements.DisplayName + ".ico")
#     }
# }
