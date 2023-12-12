#Windows:
driverquery
systeminfo
set
clip
assoc
fc
cipher
netstat -an
oing
ipconfig
sfc /scannow
powercfg /list
del
attrib +h +r +s 
start
tree
ver
tasklist
taskkill /f /pid
date
time
vol
dism
findstr
more
move
ren
shutdown /s
pathping
nslookup
chkdsk
cmdkey
makecab
mrinfo
recover
dispdiag
klist
route print

#Github:
-winget install gh (winget install github.cli)
-git auth login
- gh repo list myorgname --limit 4000 | while read -r repo _; do
  gh repo clone "$repo" "$repo"
done
