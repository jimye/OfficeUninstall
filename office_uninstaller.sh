#! /bin/sh

# Author : jim ye
# This is a shell scirpt to uninstall Microsoft Office for Mac 2011/2016/2019/365

# Reference:
# 1.https://support.microsoft.com/en-us/kb/2398768
# 2.https://support.microsoft.com/en-us/office/troubleshoot-office-for-mac-issues-by-completely-uninstalling-before-you-reinstall-ec3aa66e-6a76-451f-9d35-cba2e14e94c0?omkt=en-us&ui=en-us&rs=en-us&ad=us

delete()
{
	if [ -e "$1" ]; then
    	sudo rm -rf "$1"
    	echo "Remove $1"
    fi
}

deletefiles()
{
    for file in  "$1"*
    do
    	if [ -e "$file" ]; then
        	sudo rm -rf "$file"
        	echo "Remove $file"
        fi
    done
}

echo 'This will completely uninstall Microsoft Office for Mac 2011/2016/2019/365 from the computer'

#0 Quit all office app

echo '\n*********************************'
echo '* Start deep uninstalling:'
echo '*********************************'

#1 Remove the Binares
delete /Applications/Microsoft\ Office\ 2011
delete /Applications/Microsoft\ Communicator.app
delete /Applications/Microsoft\ Messenger.app
delete /Applications/Microsoft\ Outlook.app
delete /Applications/Microsoft\ Excel.app
delete /Applications/Microsoft\ OneNote.app
delete /Applications/Microsoft\ PowerPoint.app
delete /Applications/Microsoft\ Word.app
delete /Applications/OneDrive.app

#2 Remove the Preferences
deletefiles ~/Library/Preferences/com.microsoft
delete ~/Library/Preferences/Microsoft/Office\ 2011
deletefiles /Library/Preferences/com.microsoft
delete /Library/Preferences/com.microsoft.office.licensingV2.plist
delete /Library/LaunchDaemons/com.microsoft.office.licensing.helper.plist
delete /Library/PrivilegedHelperTools/com.microsoft.office.licensing.helper
delete /Library/LaunchDaemons/com.microsoft.office.licensingV2.helper.plist
delete /Library/LaunchDaemons/com.microsoft.autoupdate.helper.plist
delete /Library/LaunchDaemons/com.microsoft.onedriveupdaterdaemon.plist
delete /Library/PrivilegedHelperTools/com.microsoft.office.licensingV2.helper
delete /Library/PrivilegedHelperTools/com.microsoft.autoupdate.helper
delete /Library/LaunchAgents/com.microsoft.update.agent.plist

#3 Remove the Container
delete ~/Library/Containers/com.microsoft.errorreporting
delete ~/Library/Containers/com.microsoft.Excel
delete ~/Library/Containers/com.microsoft.netlib.shipassertprocess
delete ~/Library/Containers/com.microsoft.Office.setupassistant
delete ~/Library/Containers/com.microsoft.Office365ServiceV2
delete ~/Library/Containers/com.microsoft.Outlook
delete ~/Library/Containers/com.microsoft.Powerpoint
delete ~/Library/Containers/com.microsoft.RMS-XPCService
delete ~/Library/Containers/com.microsoft.Word
delete ~/Library/Containers/com.microsoft.onenote.mac
delete ~/Library/Containers/com.microsoft.onedrive.findersync
delete ~/Library/Group\ Containers/UBF8T346G9.ms
delete ~/Library/Group\ Containers/UBF8T346G9.Office
delete ~/Library/Group\ Containers/UBF8T346G9.OfficeOsfWebHost
delete ~/Library/Group\ Containers/UBF8T346G9.OfficeOneDriveSyncIntegration
delete ~/Library/Group\ Containers/UBF8T346G9.OneDriveStandaloneSuite

#4 Remove the Application Support
delete /Library/Application\ Support/Microsoft
delete ~/Library/Application\ Support/Microsoft/Office
deletefiles ~/Library/Application\ Support/CrashReporter/Microsoft\ Communicator
deletefiles ~/Library/Application\ Support/CrashReporter/Microsoft\ Messenger
deletefiles ~/Library/Application\ Support/CrashReporter/Microsoft\ Outlook
deletefiles ~/Library/Application\ Support/CrashReporter/Microsoft\ Excel
deletefiles ~/Library/Application\ Support/CrashReporter/Microsoft\ OneNote
deletefiles ~/Library/Application\ Support/CrashReporter/Microsoft\ PowerPoint
deletefiles ~/Library/Application\ Support/CrashReporter/Microsoft\ Word
deletefiles ~/Library/Application\ Support/CrashReporter/OneDrive

#5 Remove the Saved Application State
deletefiles ~/Library/Saved\ Application\ State/com.microsoft

#6 Remove the Caches
deletefiles ~/Library/Caches/com.microsoft
delete ~/Library/Caches/Microsoft
delete ~/Library/Caches/Microsoft\ Office

#7 Remove the Log
deletefiles ~/Library/Logs/DiagnosticReports/Microsoft\ Communicator
deletefiles ~/Library/Logs/DiagnosticReports/Microsoft\ Messenger
deletefiles ~/Library/Logs/DiagnosticReports/Microsoft\ Outlook
deletefiles ~/Library/Logs/DiagnosticReports/Microsoft\ Excel
deletefiles ~/Library/Logs/DiagnosticReports/Microsoft\ OneNote
deletefiles ~/Library/Logs/DiagnosticReports/Microsoft\ PowerPoint
deletefiles ~/Library/Logs/DiagnosticReports/Microsoft\ Word
deletefiles ~/Library/Logs/DiagnosticReports/OneDrive
deletefiles /Library/Logs/DiagnosticReports/Microsoft\ Communicator
deletefiles /Library/Logs/DiagnosticReports/Microsoft\ Messenger
deletefiles /Library/Logs/DiagnosticReports/Microsoft\ Outlook
deletefiles /Library/Logs/DiagnosticReports/Microsoft\ Excel
deletefiles /Library/Logs/DiagnosticReports/Microsoft\ OneNote
deletefiles /Library/Logs/DiagnosticReports/Microsoft\ PowerPoint
deletefiles /Library/Logs/DiagnosticReports/Microsoft\ Word
deletefiles /Library/Logs/DiagnosticReports/OneDrive

#8 Remove the Automater
delete /Library/Automator/Add\ Attachments\ to\ Outlook\ Messages.action
delete /Library/Automator/Add\ Content\ to\ Word\ Documents.action
delete /Library/Automator/Add\ Document\ Properties\ Page\ to\ Word\ Documents.action
delete /Library/Automator/Add\ New\ Sheet\ to\ Workbooks.action
delete /Library/Automator/Add\ Table\ of\ Contents\ to\ Word\ Documents.action
delete /Library/Automator/Add\ Watermark\ to\ Word\ Documents.action
delete /Library/Automator/Apply\ Animation\ to\ PowerPoint\ Slide\ Parts.action
delete /Library/Automator/Apply\ Font\ Format\ Settings\ to\ Word\ Documents.action
delete /Library/Automator/AutoFormat\ Data\ in\ Excel\ Workbooks.action
delete /Library/Automator/Bring\ Word\ Documents\ to\ Front.action
delete /Library/Automator/Close\ Excel\ Workbooks.action
delete /Library/Automator/Close\ Outlook\ Items.action
delete /Library/Automator/Close\ PowerPoint\ Presentations.action
delete /Library/Automator/Close\ Word\ Documents.action
delete /Library/Automator/Combine\ Excel\ Files.action
delete /Library/Automator/Combine\ PowerPoint\ Presentations.action
delete /Library/Automator/Combine\ Word\ Documents.action
delete /Library/Automator/Compare\ Word\ Documents.action
delete /Library/Automator/Convert\ Format\ of\ Excel\ Files.action
delete /Library/Automator/Convert\ Format\ of\ PowerPoint\ Presentations.action
delete /Library/Automator/Convert\ Format\ of\ Word\ Documents.action
delete /Library/Automator/Convert\ PowerPoint\ Presentations\ to\ Movies.action
delete /Library/Automator/Convert\ Word\ Content\ Object\ to\ Text\ Object.caction
delete /Library/Automator/Copy\ Excel\ Workbook\ Content\ to\ the\ Clipboard.action
delete /Library/Automator/Copy\ PowerPoint\ Slides\ to\ the\ Clipboard.action
delete /Library/Automator/Copy\ Word\ Document\ Content\ to\ the\ Clipboard.action
delete /Library/Automator/Create\ List\ from\ Data\ in\ Workbook.action
delete /Library/Automator/Create\ New\ Excel\ Workbook.action
delete /Library/Automator/Create\ New\ Outlook\ Mail\ Message.action
delete /Library/Automator/Create\ New\ PowerPoint\ Presentation.action
delete /Library/Automator/Create\ New\ Word\ Document.action
delete /Library/Automator/Create\ PowerPoint\ Picture\ Slide\ Shows.action
delete /Library/Automator/Create\ Table\ from\ Data\ in\ Workbook.action
delete /Library/Automator/Delete\ Outlook\ Items.action
delete /Library/Automator/Find\ and\ Replace\ Text\ in\ Word\ Documents.action
delete /Library/Automator/Flag\ Word\ Documents\ for\ Follow\ Up.action
delete /Library/Automator/Forward\ Outlook\ Mail\ Messages.action
delete /Library/Automator/Get\ Content\ from\ Word\ Documents.action
delete /Library/Automator/Get\ Images\ from\ PowerPoint\ Slides.action
delete /Library/Automator/Get\ Images\ from\ Word\ Documents.action
delete /Library/Automator/Get\ Parent\ Presentations\ of\ Slides.action
delete /Library/Automator/Get\ Parent\ Workbooks.action
delete /Library/Automator/Get\ Selected\ Content\ from\ Excel\ Workbooks.action
delete /Library/Automator/Get\ Selected\ Content\ from\ Word\ Documents.action
delete /Library/Automator/Get\ Selected\ Outlook\ Items.action
delete /Library/Automator/Get\ Selected\ Text\ from\ Outlook\ Items.action
delete /Library/Automator/Get\ Text\ From\ Outlook\ Mail\ Messages.action
delete /Library/Automator/Get\ Text\ from\ Word\ Documents.action
delete /Library/Automator/Import\ Text\ Files\ to\ Excel\ Workbook.action
delete /Library/Automator/Insert\ Captions\ into\ Word\ Documents.action
delete /Library/Automator/Insert\ Content\ into\ Outlook\ Mail\ Messages.action
delete /Library/Automator/Insert\ Content\ into\ Word\ Documents.action
delete /Library/Automator/Insert\ New\ PowerPoint\ Slides.action
delete /Library/Automator/Mark\ Outlook\ Mail\ Message\ as\ a\ To\ Do\ Item.action
delete /Library/Automator/Office.definition
delete /Library/Automator/Open\ Excel\ Workbooks.action
delete /Library/Automator/Open\ Outlook\ Items.action
delete /Library/Automator/Open\ PowerPoint\ Presentations.action
delete /Library/Automator/Open\ Word\ Documents.action
delete /Library/Automator/Paste\ Clipboard\ Content\ into\ Excel\ Workbooks.action
delete /Library/Automator/Paste\ Clipboard\ Content\ into\ Outlook\ Items.action
delete /Library/Automator/Paste\ Clipboard\ Content\ into\ PowerPoint\ Presentations.action
delete /Library/Automator/Paste\ Clipboard\ Content\ into\ Word\ Documents.action
delete /Library/Automator/Play\ PowerPoint\ Slide\ Shows.action
delete /Library/Automator/Print\ Excel\ Workbooks.action
delete /Library/Automator/Print\ Outlook\ Messages.action
delete /Library/Automator/Print\ PowerPoint\ Presentations.action
delete /Library/Automator/Print\ Word\ Documents.action
delete /Library/Automator/Protect\ Word\ Documents.action
delete /Library/Automator/Quit\ Excel.action
delete /Library/Automator/Quit\ Outlook.action
delete /Library/Automator/Quit\ PowerPoint.action
delete /Library/Automator/Quit\ Word.action
delete /Library/Automator/Reply\ to\ Outlook\ Mail\ Messages.action
delete /Library/Automator/Save\ Excel\ Workbooks.action
delete /Library/Automator/Save\ Outlook\ Draft\ Messages.action
delete /Library/Automator/Save\ Outlook\ Items\ as\ Files.action
delete /Library/Automator/Save\ Outlook\ Messages\ as\ Files.action
delete /Library/Automator/Save\ PowerPoint\ Presentations.action
delete /Library/Automator/Save\ Word\ Documents.action
delete /Library/Automator/Search\ Outlook\ Items.action
delete /Library/Automator/Select\ Cells\ in\ Excel\ Workbooks.action
delete /Library/Automator/Select\ PowerPoint\ Slides.action
delete /Library/Automator/Send\ Outgoing\ Outlook\ Mail\ Messages.action
delete /Library/Automator/Set\ Category\ of\ Outlook\ Items.action
delete /Library/Automator/Set\ Document\ Settings.action
delete /Library/Automator/Set\ Excel\ Workbook\ Properties.action
delete /Library/Automator/Set\ Footer\ for\ PowerPoint\ Slides.action
delete /Library/Automator/Set\ Outlook\ Contact\ Properties.action
delete /Library/Automator/Set\ PowerPoint\ Slide\ Layout.action
delete /Library/Automator/Set\ PowerPoint\ Slide\ Transition\ Settings.action
delete /Library/Automator/Set\ Security\ Options\ for\ Word\ Documents.action
delete /Library/Automator/Set\ Text\ Case\ in\ Word\ Documents.action
delete /Library/Automator/Set\ Word\ Document\ Properties.action
delete /Library/Automator/Sort\ Data\ in\ Excel\ Workbooks.action

#9 Remove the license file
delete /Library/Preferences/com.microsoft.office.licensing.plist
delete /Library/Preferences/com.microsoft.office.licensingV2.plist
deletefiles ~/Library/Preferences/ByHost/com.microsoft

#10 Remove the receipts
deletefiles /Library/Receipts/Office2011_
deletefiles /Library/Receipts/Office2016_
deletefiles /Library/Receipts/Office2019_
deletefiles /private/var/db/receipts/com.microsoft.office

#11 Remove the Microsoft fonts
delete /Library/Fonts/Microsoft

#12 Remove the Microsoft ~ Data folder
delete ~/Documents/Microsoft\ ~\ Data

#13 Remove Internet Plug-Ins
deletefiles /Library/Internet\ Plug-Ins/SharePoint

#14 Remove cookies
delete ~/Library/Cookies/com.microsoft.onedrive.binarycookies
delete ~/Library/Cookies/com.microsoft.onedriveupdater.binarycookies

echo '\nDone\n'
echo '****************************************************'
echo '*                   Congratulations!               *'
echo '*                                                  *'
echo '*      All files of Office for mac are removed!    *'
echo '****************************************************'

echo '\n**********************************************************'
echo '* Now, follow these steps to finish completely uninstall:'
echo '**********************************************************'
echo '1 Please remove Keychain Entries:'
echo 'Open Finder > Applications > Utilities > Keychain Access and remove the following password entries:\n'
echo '\tMicrosoft Office Identities Cache 2'
echo '\tMicrosoft Office Identities Settings 2\n'

echo '2 Search for all occurrences of ADAL in the keychain and remove all those entries if present.\n'

echo '3 Remove Office for Mac icons from the Dock'
echo 'If you added Office icons to the Dock they may turn into question marks after you uninstall Office for Mac. To remove these icons, control+click or right-click the icon and click Options > Remove from Dock.\n'

echo '4 Restart your computer'
