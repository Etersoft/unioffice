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

#define VER_2 1
#define VER_3 2

extern LONG g_cServerLocks;
extern LONG g_cComponents;

extern LONG OOVersion;








#endif //__UNIOFFICE_EXCEL_PRIVATE_H__
