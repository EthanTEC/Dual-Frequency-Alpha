# How to create an update:

* Ensure the UPDATE_INFO_URL is pointed to the raw update_info.json string kept in main, this will house your version and direct the updater to the correct update

* Ensure any changes made to the python file have been pushed (doesn't effect the update creation process just good practice)

* In a terminal in your repo folder (I'm on windows so yours might be slightly different) run "pyinstaller --noconfirm --onefile --windowed --add-data "Images;Images" Python/AlphaAnalysisApp.py"

* This will create a couple folders, the one we care about is the dist folder, which should holder the AlphaAnalysisApp.exe file

* Create a new release and tag it with the version of the update (ie v1.2.3)

* Upload the exe file to the patch, and press upload once it finishes loading

* Publish the release and copy the URL of the release to your clipboard (Something like "https://github.com/user/repo/releases/download/v1.0.1/patch)

* Paste the new link into update_info.json under "download_url" so the code knows where to find the most recent version

* Push this so the hosting link saved in the python code now sees the new patch url

* Now when a user opens the app or presses "Check for Updates" the app will check the json for the location of the most recent update, compare version with it, and if the app is on an older verions, replaces it
