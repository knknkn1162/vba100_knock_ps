# vba100_knock (Powershell version)

+ [INFO] This is under development..

+ Powershell version of [knknkn1162/vba100_knock](https://github.com/knknkn1162/vba100_knock). See also https://excel-ubara.com/vba100/.

# requirements

+ Windows >= 10
+ Powershell >= 5.1

# Prerequisites

+ Install chocolatey, make
+ (Optional) ghostscript, imagemagick.app

```ps
# scripts to be runnable
Set-ExecutionPolicy RemoteSigned
# install commands in Admin
Start-Process powershell -Verb runAs
choco install -y make

# (Optional) when capture
Start-Process powershell -Verb runAs
choco install -y imagemagick.app ghostscript
## specify version
$ENV:Path="C:\Program Files\ImageMagick-${version};"+$ENV:Path
```

# How to run scripts

```ps
# 1. (optional) create shell script from template
make template XLSM=ex001
# 2. Edit your code
# 3. run the script as macro
make run XLSM=ex001
# make run XLSM=ex001 DEBUG=0 # save after macro, run faster
```
