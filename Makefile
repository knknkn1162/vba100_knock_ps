SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -ExecutionPolicy RemoteSigned -Command

BOOKS_DIR=books
SRC_DIR=src
MAIN_PS1=$(SRC_DIR)/main.ps1

# `make <action> XLSM=ex008
XLSM_BASENAME=$(XLSM)
XLSM_NAME=$(XLSM).xlsm
XLSM_RELPATH=$(BOOKS_DIR)/$(XLSM_NAME)
XLSM_ABSPATH=$(abspath $(XLSM_RELPATH))
SCRIPT_NAME=$(XLSM).ps1
SCRIPT_PATH=$(abspath $(SRC_DIR)/$(SCRIPT_NAME))

DEBUG=1
DEBUG_OPTION=-debug
ifeq ($(DEBUG), 0)
	DEBUG_OPTION=
endif

.PHONY: run
run:
	$(MAIN_PS1) -pspath $(SCRIPT_PATH) -xlpath $(XLSM_ABSPATH) $(DEBUG_OPTION)
