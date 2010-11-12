
include ../Makefile.mingw

#IDL=gdb --args /srv/wine/Projects/wine-origin/tools/widl/widl
IDL=widl

# TODO: remove msvcrt using (do not use functions from it)
LIBS=-ladvapi32 -lole32 -loleaut32 -luser32 -lkernel32 -lshell32 -luuid -lshlwapi -lmsvcrt

OBJECTS = unioffice_excel_i.o \
	tools.o \
	main.o \
	factory.o \
	regsvr.o \
	i_border.o \
	i_borders.o \
	i_interior.o \
	i_font.o  \
	special_functions.o \
	i_workbooks.o \
	i_names.o \
	i_workbook.o \
	i_range.o \
	i_pagesetup.o\
	i_shape.o \
	i_shapes.o \
	i_sheets.o \
	i_outline.o\
	i_worksheet.o \
	i_windows.o \
	application_excel.o

.PHONY: all clean

all: unioffice_excel.dll

unioffice_excel.dll: $(OBJECTS)
	$(CC) --add-stdcall-alias -shared -o $@ $^ $(LIBS)
	$(STRIP) --strip-all $@
#	$(CC) -shared -o $@ $^ $(LDADD)

# needed wine-devel
unioffice_excel_i.c:
	$(IDL) -I /usr/include/wine/windows -u -h -t unioffice_excel.idl

clean:
	rm -f *.o *.dll *.h~ *.c~ *.idl~ unioffice_excel.h unioffice_excel_i.c *.tlb
