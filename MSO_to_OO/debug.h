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

#define DEBUG 1
#ifdef DEBUG
#define TRACE(args...) \
do { fprintf(stderr,"%s:%s:",__FILE__,__FUNCTION__);fprintf(stderr, args); } while(0)
#else
#define TRACE(n,...)
#endif

#ifdef DEBUG
#define WTRACE(args...) \
do { fprintf(stderr,"%s:%s:",__FILE__,__FUNCTION__);fwprintf(stderr, args); } while(0)
#else
#define WTRACE(n,...)
#endif

