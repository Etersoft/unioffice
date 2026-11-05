include ../Makefile.mingw

MODULE = unioffice_excel
LINK = $(CC)
LIBS = -ladvapi32 -lole32 -loleaut32 -luser32 -lkernel32 -lshell32 -luuid -lshlwapi -lmsvcrt

OBJECT_NAMES = unioffice_excel_i.o \
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

include ../Makefile.component
