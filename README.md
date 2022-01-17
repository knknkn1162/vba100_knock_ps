# vba100_knock (Powershell version)

+ [INFO] This is under development..

+ Powershell version of [knknkn1162/vba100_knock](https://github.com/knknkn1162/vba100_knock). See also https://excel-ubara.com/vba100/.

# requirements

+ Windows >= 10
+ Powershell >= 5.1

# Prerequisites

+ Install chocolatey, nkf, make
+ (Optional) ghostscript, imagemagick.app

```ps
# scripts to be runnable
Set-ExecutionPolicy RemoteSigned
# install commands in Admin
Start-Process powershell -Verb runAs
choco source add -n kai2nenobu -s https://www.myget.org/F/kai2nenobu
choco install -y nkf make

# (Optional) when capture
Start-Process powershell -Verb runAs
choco install -y imagemagick.app ghostscript
## specify version
$ENV:Path="C:\Program Files\ImageMagick-${version};"+$ENV:Path
```

# How to run

```ps
# specify basename of xlsm file.
make run XLSM=ex001
```
