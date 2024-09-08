```powershell
# 1. Check if PS profile exists
Test-Path $PROFILE

# 1.a If not found (returns "False") run below command to create PS profile
New-Item -ItemType File -Path $PROFILE

# 1.b If not found (returns "False") run below command to create scripts folder
New-Item -ItemType Directory -Path $(Split-Path $PROFILE) -Name Scripts

# 2. Open the profile, can use any editor
notepad.exe $PROFILE

# 3. Add the comment and line to profile
# Import local script(s)
. "$(Split-Path -Path $PROFILE)\Scripts\bttc-cli.ps1"

# 4. Add the cli script file to your Scripts folder where the profile is

# 5. Right click and go to Properties -> General Tab, and check the "Unblock" at the bottom to enable the script, by default Windows blocks execution and then you would need to open a new powershell instance
```