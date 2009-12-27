/*
 * Main header file
 *
 * Copyright (C) 2009 Sinitsin Ivan (Etersoft) <ivan@etersoft.ru>
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

#define _WIN32_WINNT 0x0600

#ifndef __UNIOFFICE_EXCEL_PRIVATE_H__
#define __UNIOFFICE_EXCEL_PRIVATE_H__

#include <stdarg.h>

#include <windows.h>
#include <windef.h>
#include <winbase.h>
#include <winuser.h>
#include <winreg.h>
#include <ole2.h>
#include <ocidl.h>
#include <oaidl.h>
#include <stddef.h>
#include "unioffice_excel.h"

#include "../Common/debug.h"
#include "../Common/tools.h"
#include "../Common/special_functions.h"
#include "../OOWrappers/oo_constants.h"

#define VER_2 1
#define VER_3 2

extern LONG g_cServerLocks;
extern LONG g_cComponents;

extern LONG OOVersion;

static long color[56] = {
    0x000000, 0xFFFFFF, 
	0xFF0000, 0x00FF00, 
	0x0000FF, 0xFFFF00, 
	0xFF00FF, 0x00FFFF, 
	0x800000, 0x008000, 
	0x000080, 0x808000, 
	0x800080, 0x008080, 
	0xC0C0C0, 0x808080, 
	0x9999FF, 0x993366, 
	0xFFFFCC, 0xCCFFFF, 
	0x660066, 0xFF8080, 
	0x0066CC, 0xCCCCFF, 
	0x000080, 0xFF00FF, 
	0xFFFF00, 0x00FFFF, 
	0x800080, 0x800000, 
	0x008080, 0x0000FF, 
	0x00CCFF, 0xCCFFFF, 
	0xCCFFCC, 0xFFFF99, 
	0x99CCFF, 0xFF99CC, 
	0xCC99FF, 0xFFCC99, 
	0x3366FF, 0x33CCCC, 
	0x99CC00, 0xFFCC00, 
	0xFF9900, 0xFF6600, 
	0x666699, 0x969696, 
	0x003366, 0x339966, 
	0x003300, 0x333300, 
	0x993300, 0x993366, 
	0x333399, 0x333333
	};

const int underline_style_NONE           = 0;
const int underline_style_SINGLE         = 1;
const int underline_style_DOUBLE         = 2;
const int underline_style_DOTTED         = 3;
const int underline_style_DONTKNOW       = 4;
const int underline_style_DASH           = 5;
const int underline_style_LONGDASH       = 6;
const int underline_style_DASHDOT        = 7;
const int underline_style_DASHDOTDOT     = 8;
const int underline_style_SMALLWAVE      = 9;
const int underline_style_WAVE           = 10;
const int underline_style_DOUBLEWAVE     = 11;
const int underline_style_BOLD           = 12;
const int underline_style_BOLDDOTTED     = 13;
const int underline_style_BOLDDASH       = 14;
const int underline_style_BOLDLONGDASH   = 15;
const int underline_style_BOLDDASHDOT    = 16;
const int underline_style_BOLDDASHDOTDOT = 17;
const int underline_style_BOLDWAVE       = 18;





#endif //__UNIOFFICE_EXCEL_PRIVATE_H__
