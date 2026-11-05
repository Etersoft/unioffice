EXCEL_BUILD ?= ../unioffice_excel/build/mingw-32
OUTPUT ?= build/unioffice.msi

.PHONY: all clean
all:
	python3 ../scripts/build-msi.py --excel-build "$(EXCEL_BUILD)" --output "$(OUTPUT)"

clean:
	rm -f "$(OUTPUT)"
