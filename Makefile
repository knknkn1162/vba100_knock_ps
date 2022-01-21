SHELL:=powershell.exe
.SHELLFLAGS:= -NoProfile -ExecutionPolicy RemoteSigned -Command

THIS_ENCODING=UTF8
# shift_jis(check with [Text.Encoding]::Default.WebName)
EXCEL_ENCODING=default
BOOKS_DIR=books
SRC_DIR=src
SRC_ENC_DIR=src_enc
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
ENC_SCRIPT_PATH=$(abspath $(SRC_ENC_DIR)/$(SCRIPT_NAME))

COMMIT_MSG="implement"

.PHONY: run template
run: create-$(SRC_ENC_DIR)
	gc -en $(THIS_ENCODING) $(SCRIPT_PATH) | Out-File -en $(EXCEL_ENCODING) $(ENC_SCRIPT_PATH)
	$(MAIN_PS1) -pspath $(ENC_SCRIPT_PATH) -xlpath $(XLSM_ABSPATH) $(DEBUG_OPTION)

template:
	cp ./template/template.ps1 $(SCRIPT_PATH)

.PHONY: push commit clean
push: commit
	git push
commit:
	git add $(SCRIPT_PATH)
	git commit -m "$(COMMIT_MSG) ps in $(XLSM_NAME)"

clean:
	rm -r -fo $(SRC_ENC_DIR)

create-$(SRC_ENC_DIR):
	if (!(Test-Path $(SRC_ENC_DIR) )) { mkdir $(SRC_ENC_DIR) }
