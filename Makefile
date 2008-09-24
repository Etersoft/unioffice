
.PHONY: all clean

all:
#	$(MAKE) -C dll
#	$(MAKE) -C test
	$(MAKE) -C unioffice_excel
#	$(MAKE) -C test_MSO
	$(MAKE) -C MSI

clean:
#	$(MAKE) clean -C dll
#	$(MAKE) clean -C test
	$(MAKE) clean -C unioffice_excel
#	$(MAKE) clean -C test_MSO
	$(MAKE) clean -C MSI