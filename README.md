# Sys_InputSender  
## Sending keystrokes, mouse-moves or mouse-clicks to the active window  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Sys_InputSender?style=plastic)](https://github.com/OlimilO1402/Sys_InputSender/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Sys_InputSender?style=plastic)](https://github.com/OlimilO1402/Sys_InputSender/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Sys_InputSender/total.svg)](https://github.com/OlimilO1402/Sys_InputSender/releases/download/v2025.5.14/InputSender_v2025.5.14.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)



Project started in 2008.  

WinAPI SendInput
----------------
This example shows how to use the Windows-API function [SendInput](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput) for sending keystrokes or mouseevents to the active window.  
This function uses an array of [Input structure](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-input) which is in general a union made of 3 different structures:
* [MOUSEINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput)
* [KEYBDINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-keybdinput)
* [HARDWAREINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-hardwareinput)

Hardwareinput
-------------
The informations about hardwareinput provided by Microsoft are useless, so I assume that maybe once it was meant for joystick-inputs under Win9x which are no longer supported in newer Windows versions.  
But why did Microsoft not glued a deprecated badge onto it long ago? I am not 100% sure about the hardwareinput-situation, and so I will leave it there as it is.

Structure Input
---------------
Of course the Array of Input must be a contiguous piece of memory, and so does a ordinary array as ud-type in VB, it uses a contiguous block of memory either.
There is no problem using the SendInput-fuction in VB, right? Well - and what about the union-type?

Union Input
-----------
You may wonder anyway what the heck is a "union"? A Union is a datatype which uses the same memory for different datatypes of different meaning and purposes, btw. just like the VB intrinsic datatype Variant does.
When we want to make use of SendInput in a typical fully object oriented windows desktop-app, of course the user wants to create and also edit all the data.
To deal with the union there are different approaches:
* we could copy all data in and out the array
* we could leave the data in the memory block and use a pointer instead
Of course in C we would use a pointer to a Input-struct. The VB-Project Sys_InputSender does it pretty much in the same manor, it uses the udt-pointer method to create and edit the data.

The udt-pointer method for VB
-----------------------------
After playing around with SafeArrays I figured out how to use it with ud-types, and shared it with other VB-coders at ActiveVB in 2008.  
How does the udt-pointer method work? We copy the pointer to a SafeArray-descriptor into an array of ud-type, and set the pvData-pointer where we want it to point to
We create a class for every Input-structure MOUSEINPUT KEYBDINPUT and HARDWAREINPUT. Every class holds a SafeArray-descriptor and a variable of its input-type
During creation of the object the pvDate of the SafeArray-descriptor points to the internal input-variable. When we add the object to the Array-List, pvData will be set to point to the variable in the array and the data will be copied to the array, with just a private udtype-assignment inside the class. It is even not necessary to get access to the Array in the List itself. The object now gets freed, as we no longer need it.

![InputSender Image](Resources/InputSender.png "InputSender Image")
