/*
 * Debug functions
 *
 * Copyright (C) 2008 Sinitsin Ivan (Etersoft) <ivan@etersoft.ru>
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301, USA
 */
#include <stdio.h>

extern char buf[MAX_PATH+50];
extern write_log;
extern FILE *trace_file;

#define DEBUG

#ifdef DEBUG
#define TRACE(args...) \
do { if (write_log) { \
trace_file = fopen(buf,"a");\
if (trace_file) { \
fprintf(trace_file,"%s:%s:",__FILE__,__FUNCTION__);\
fprintf(trace_file, args); \
if (trace_file) fclose(trace_file);\
} \
} \
} while(0)
#else
#define TRACE(n,...)
#endif

#ifdef DEBUG
#define WTRACE(args...) \
do { if (write_log) { \
trace_file = fopen(buf,"a");\
if (trace_file) { \
fprintf(trace_file,"%s:%s:",__FILE__,__FUNCTION__);\
fwprintf(trace_file, args); \
if (trace_file) fclose(trace_file);\
} \
} \
} while(0)
#else
#define WTRACE(n,...)
#endif

#define TRACE_IN TRACE(" IN \n")
#define TRACE_OUT TRACE(" OUT \n")
#define TRACE_NOTIMPL TRACE("ERROR method not implement \n")

#define CREATE_OBJECT TRACE("Create object \n")
#define DELETE_OBJECT TRACE("Delete object \n")


#ifdef DEBUG
#define ERR(args...) \
do { if (write_log) { \
trace_file = fopen(buf,"a");\
if (trace_file) { \
fprintf(trace_file,"ERROR:%s:%s:",__FILE__,__FUNCTION__);\
fprintf(trace_file, args); \
if (trace_file) fclose(trace_file);\
} \
} \
} while(0)
#else
#define ERR(n,...)
#endif

#ifdef DEBUG
#define WERR(args...) \
do { if (write_log) { \
trace_file = fopen(buf,"a");\
if (trace_file) { \
fprintf(trace_file,"ERROR:%s:%s:",__FILE__,__FUNCTION__);\
fwprintf(trace_file, args); \
if (trace_file) fclose(trace_file);\
} \
} \
} while(0)
#else
#define WERR(n,...)
#endif


