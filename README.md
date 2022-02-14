# vba100_knock (Powershell version)

+ Powershell version of [knknkn1162/vba100_knock](https://github.com/knknkn1162/vba100_knock). See also https://excel-ubara.com/vba100/.

## note

+ we skip several exercises because of the following difficuties:
    + ex054: worksheet event
    + ex055: worksheet event
    + ex068: form control
    + ex070: timer event
    + ex073: form control
    + ex076: form control
    + ex077: worksheet event
    + ex080: worksheet event

# requirements

+ Windows >= 10
+ Powershell >= 5.1

# Prerequisites

+ Install chocolatey, make
+ (Optional) ghostscript, imagemagick.app
+ (for ex100) AngleParse

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

# (For ex100) See https://github.com/kamome283/AngleParse
Install-Module AngleParse
```

# How to run scripts

```ps
# 1. (optional) create shell script from template
make template XLSM=ex001
# 2. Edit your code
# 3. run the script as macro
make run XLSM=ex001
# make run XLSM=ex001 DEBUG=0 # save after macro, run faster
# 4. If you want to cleanup and initialize dirty outputs and inputs, `make clean`
make clean XLSM=ex057
```
