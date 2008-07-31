
.PHONY: all clean

all:
#	$(MAKE) -C dll
#	$(MAKE) -C test
	$(MAKE) -C MSO_to_OO
#	$(MAKE) -C test_MSO
	$(MAKE) -C MSI

clean:
#	$(MAKE) clean -C dll
#	$(MAKE) clean -C test
	$(MAKE) clean -C MSO_to_OO
#	$(MAKE) clean -C test_MSO
	$(MAKE) clean -C MSI