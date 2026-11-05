.DEFAULT_GOAL := all
include Makefile.mingw

.PHONY: all libraries excel word msi clean
all: libraries
libraries: excel word
excel:
	$(MAKE) -C unioffice_excel
word:
	$(MAKE) -C unioffice_word
msi: libraries
	$(MAKE) -C MSI EXCEL_BUILD=../unioffice_excel/$(BUILD_DIR) OUTPUT=build/$(TOOLCHAIN)-$(BITS)/unioffice.msi
clean:
	$(MAKE) clean -C unioffice_excel
	$(MAKE) clean -C unioffice_word
