```powershell
#Check if profile exists
Test-Path $PROFILE

#If not run command to create profile
New-Item -ItemType File -Path $PROFILE

#Open the profile
notepad.exe $PROFILE

# Add the comment and line to profile
# Import local scripts
. "$(Split-Path -Path $PROFILE.CurrentUserCurrentHost)\Scripts\bttc-cli.ps1"

# Add the cli script file to your Scripts folder where the profile is
```