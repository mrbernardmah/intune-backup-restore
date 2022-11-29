::Name - Restore Script
::Description - Restores Up favourites, signatures, Outlook settings, templates, UI settings, stickynotes, chrome bookmarks, email signature and mapped drives from OneDrive to local machine


::GET FAVOURITES
set error="%OneDriveCommercial%\backup\restoreerrorlog.txt" 

echo CopyingFavourites >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BFav=%USERPROFILE%\Favorites
SET RFav=%OneDriveCommercial%\Favorites
XCopy "%RFav%\*" "%BFav%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 

::GET SIGNATURES
echo CopyingSignatures >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BSig=%APPDATA%\Microsoft\Signatures
SET RSig=%OneDriveCommercial%\Backup\Signature
if not exist "%APPDATA%\Microsoft\Signatures" mkdir "%APPDATA%\Microsoft\Signatures"
XCopy "%RSig%\*" "%BSig%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 


::GET Outlook-Autocorrect
echo CopyingOutlookComplete >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BAuto=%LOCALAPPDATA%\Microsoft\Outlook\RoamCache
SET RAuto=%OneDriveCommercial%\Backup\Roam
if not exist "%LOCALAPPDATA%\Microsoft\Outlook\RoamCache" mkdir "%LOCALAPPDATA%\Microsoft\Outlook\RoamCache"
XCopy "%RAuto%\*" "%BAuto%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 


::GET NORMALDOT
echo CopyingNormalDot >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BWord=%APPDATA%\Microsoft\Templates
SET RWord=%OneDriveCommercial%\Backup\Templates
if not exist "%APPDATA%\Microsoft\Templates" mkdir "%APPDATA%\Microsoft\Templates"
XCopy "%RWord%\*" "%BWord%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 

::GET Links
echo CopyingLinks >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BLinks=B%USERPROFILE%\Links
SET RLinks=%OneDriveCommercial%\Backup\Links
if not exist "%APPDATA%\Microsoft\Templates" mkdir "%APPDATA%\Microsoft\Templates"
XCopy "%RWord%\*" "%BWord%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 


::GET ExcelStart
echo CopyingExcelStart >> "%OneDriveCommercial%\backup\restorelog.txt"
SET Bxlstart=%APPDATA%\Microsoft\Excel\XLStart
SET Rxlstart=%OneDriveCommercial%\Backup\XLStart
if not exist "%APPDATA%\Microsoft\Excel\XLStart" mkdir "%APPDATA%\Microsoft\Excel\XLStart"
XCopy "%Rxlstart%\*" "%Bxlstart%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 

::GET WordStart
echo CopyingWordStart >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BWordst=%APPDATA%\Microsoft\Word\Startup
SET RWordst=%OneDriveCommercial%\Backup\Wordstartup
if not exist "%APPDATA%\Microsoft\Word\Startup" mkdir "%APPDATA%\Microsoft\Word\Startup"
XCopy "%RWordst%\*" "%BWordst%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 

::GET UI
echo CopyingUI >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BUI=%LOCALAPPDATA%\Microsoft\Office
SET RUI=%OneDriveCommercial%\Backup\UI
if not exist "%LOCALAPPDATA%\Microsoft\Office" mkdir "%LOCALAPPDATA%\Microsoft\Office"
XCopy "%RUI%\*.customUI" "%BUI%" /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 
XCopy "%RUI%\*.officeUI" "%BUI%" /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 


::GET SICKYNOTES
echo CopyingStickyNotes >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BSticky=%LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\legacy
SET RSticky=%OneDriveCommercial%\Backup\Sticky
if not exist "%LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\legacy" mkdir "%LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\legacy"
XCopy /s "%RSticky%" "%BSticky%" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 
copy "%BSticky%\StickyNotes.snt" "%BSticky%\ThresholdNotes.snt"

::GET CHROME BOOKMARKS
echo CopyingChromeBookmarks >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BChrome=%LOCALAPPDATA%\Google\Chrome\User Data\Default
SET RChrome=%OneDriveCommercial%\Backup\Chrome
if not exist "%OneDriveCommercial%\Backup\Chrome" mkdir "%OneDriveCommercial%\Backup\Chrome"
XCopy "%RChrome%\Bookmarks*" "%BChrome%\Bookmarks*" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 

::GET EDGE BOOKMARKS
echo CopyingEdgeookmarks >> "%OneDriveCommercial%\backup\restorelog.txt"
SET BEdge=%LOCALAPPDATA%\Microsoft\Edge\User Data\Default
SET REdge=%OneDriveCommercial%\Backup\Edge
if not exist "%Onedrive%\Backup\Edge" mkdir "%Onedrive%\Backup\Edge"
XCopy "%BEdge%\Bookmarks*" "%BEdge%\Bookmarks*" /E /Y /C /Z /D >> "%OneDriveCommercial%\backup\restorelog.txt" 2>> %error% 

::Set Signature
FOR %%f IN (%APPDATA%\Microsoft\Signatures\*.htm) DO (
 set filename=%%~nf
 goto tests
)
:tests
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\XX.X\Common\MailSettings\16.0\Common\MailSettings" /v NewSignature /t REG_EXPAND_SZ /d %filename%
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\XX.X\Common\MailSettings\16.0\Common\MailSettings" /v ReplySignature /t REG_EXPAND_SZ /d %filename%

md C:\Restore1
cmd /c xcopy Restore1.txt C:\Restore1