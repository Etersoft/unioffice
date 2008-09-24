
.PHONY: all clean

all:
	$(MAKE) -C unioffice_excel
	$(MAKE) -C unioffice_word
	$(MAKE) -C MSI

clean:
	$(MAKE) clean -C unioffice_excel
	$(MAKE) clean -C unioffice_word
	$(MAKE) clean -C MSI