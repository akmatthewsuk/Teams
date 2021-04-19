:: Clean output directories
rmdir obj\Debug /s /q
rmdir bin\Debug /s /q

:: Compile folders
"%Wix%\bin\heat.exe" dir .\Images -cg Images -ag -srd -dr IMAGESFOLDER -out Images.wxs
"%Wix%\bin\heat.exe" dir .\Scripts -cg Scripts -ag -srd -dr SCRIPTSFOLDER -out SCRIPTS.wxs

:: Make the MSI
"%Wix%\bin\candle.exe" -out obj\Debug\ -arch x64 -ext .\PowerShellWixExtension.2.0.1\tools\lib\PowerShellWixExtension.dll Product.wxs Images.wxs SCRIPTS.wxs
"%Wix%\bin\light.exe" -out bin\Debug\TeamsBackgrounds.msi -ext .\PowerShellWixExtension.2.0.1\tools\lib\PowerShellWixExtension.dll obj\Debug\Product.wixobj obj\debug\Images.wixobj obj\debug\Scripts.wixobj -b Images -b Scripts
