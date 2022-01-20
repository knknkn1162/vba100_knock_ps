SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -ExecutionPolicy RemoteSigned -Command

BOOKS_DIR=books
SRC_DIR=src
MAIN_PS1=$(SRC_DIR)/main.ps1

ifneq ("$(OS)", "Windows_NT")
$(warning [WARNING] COM only works on Windows machine.)
endif
ifeq (,$(XLSM))
$(error XLSM variable is not set!)
endif
DEBUG=1
DEBUG_OPTION=-debug
ifeq ($(DEBUG), 0)
	DEBUG_OPTION=
endif

# `make <action> XLSM=ex008
XLSM_BASENAME=$(XLSM)
XLSM_NAME=$(XLSM).xlsm
XLSM_RELPATH=$(BOOKS_DIR)/$(XLSM_NAME)
XLSM_ABSPATH=$(abspath $(XLSM_RELPATH))
SCRIPT_NAME=$(XLSM).ps1
SCRIPT_PATH=$(abspath $(SRC_DIR)/$(SCRIPT_NAME))

COMMIT_MSG="implement"

.PHONY: run
run:
	$(MAIN_PS1) -pspath $(SCRIPT_PATH) -xlpath $(XLSM_ABSPATH) $(DEBUG_OPTION)

push: commit
	git push
commit:
	git add $(SCRIPT_PATH)
	git commit -m "$(COMMIT_MSG) ps in $(XLSM).xlsm"
