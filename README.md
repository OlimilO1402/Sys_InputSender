# Sys_InputSender  
## Sending keystrokes, mouse-moves or mouse-clicks to the active window  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Sys_InputSender?style=plastic)](https://github.com/OlimilO1402/Sys_InputSender/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Sys_InputSender?style=plastic)](https://github.com/OlimilO1402/Sys_InputSender/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Sys_InputSender/total.svg)](https://github.com/OlimilO1402/Sys_InputSender/releases/download/v2025.5.14/InputSender_v2025.5.14.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)



Project started in 2008.  

What is it
----------
This example shows how to use the Windows-API function [SendInput](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput) for sending keystrokes or mouseevents to the active window. 
This function uses an array of [struct INPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-input) which is in general a union made of 3 different structures:
* [MOUSEINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput)
* [KEYBDINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-keybdinput)
* [HARDWAREINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-hardwareinput)

MOUSEINPUT & KEYBDINPUT
-----------------------
With MOUSEINPUT you simulate mouse-movemements and mouse-clicks. With KEYBDINPUT you simulate Key-down events and key-up events. 
But be careful, you have to know what you do, because with it you can easily screw up your system, and it may behave weird.
After a key-down of any key, a key-up must follow. e.g.  
Shift-key-down + A-Key-down + A-Key-up + Shift-Key-up: result is the upper case letter "A".

HARDWAREINPUT
-------------
The informations about hardwareinput provided by learn.Microsoft are pretty much useless. We must assume that maybe once it was meant for joystick-inputs or other hardware inputs under Win9x which are no longer supported in newer Windows versions.  
But why did Microsoft not glued a deprecated-badge onto it long ago? I am not 100% sure about the hardwareinput-situation, and so I will leave it there as it is.

WinAPI SendInput
----------------
The array of struct INPUT must be a contiguous piece of memory, and so does a ordinary array as ud-type in VB, it uses a contiguous block of memory either.
We consider a ArrayList class for managing the memory and adding or deleting elements, lets call it WndInputs.
So there is no problem using the SendInput-fuction in VB, right? Well - but what about the union type?

The Union in struct INPUT
-------------------------
You may wonder anyway what the heck is a "union"? A union is a datatype which uses the same memory for different datatypes of different meaning and purposes, btw. just like the VB intrinsic datatype Variant does. Other than Variant, VB is not capable of creating a union type out of the box, we have to use a little bit of a trick.
In a typical fully object oriented windows desktop app, of course the user wants to create and change and edit or delete the data.
When we want to make use of the SendInput function in a fully editable way we must find a way to use the union.  
To deal with the array of struct and union in VB we could use different approaches:
a) we could collect all data anywhere in the heap memory, and at the time when SendInput is called we copy all data to an array
b) we could copy all data in and out the array for every editing
c) we could leave the data in the memory block and use a pointer instead  

Pointer to a Structure
----------------------
Of course in C we would use a pointer to the struct INPUT. This repo Sys_InputSender does it pretty much in the same manor, it uses the udt-pointer method to create and edit the data.

The udt-pointer method for VB
-----------------------------
After playing around with SafeArrays I figured out how to use it with ud-types, and shared it with other VB-coders at ActiveVB in 2008.  
How does the udt-pointer method work? We copy the pointer to a SafeArray-descriptor into an empty array of ud-type, and set the pvData-pointer where we want it to point to. 
We create a class for every Input-structure MOUSEINPUT KEYBDINPUT and HARDWAREINPUT. Every class holds a SafeArray-descriptor and a variable of its input-type.
During creation of the object the pvDate of the SafeArray-descriptor points to the internal input-variable. When we add the object to the Array-List, pvData will be set to point to the variable in the array and the data will be copied to the array, with just a private udtype-assignment inside the class. It is even not necessary to get access to the Array in the List itself. The object now gets freed, as we no longer need it.

![InputSender Image](Resources/InputSender.png "InputSender Image")
