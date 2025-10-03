# Sys_InputSender  
## Sending keystrokes, mouse-moves or mouse-clicks to the active window  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Sys_InputSender?style=plastic)](https://github.com/OlimilO1402/Sys_InputSender/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Sys_InputSender?style=plastic)](https://github.com/OlimilO1402/Sys_InputSender/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Sys_InputSender/total.svg)](https://github.com/OlimilO1402/Sys_InputSender/releases/download/v2025.10.3/InputSender_v2025.10.3.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)



Project started in 2008.  

What is it
----------
This project shows how to use the Windows-API function [SendInput](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendinput) for sending keystrokes or mouseevents to the active window. 
It uses an array of [INPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-input) structure, which is basically a union made of 3 different structures:
* [MOUSEINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput)
* [KEYBDINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-keybdinput)
* [HARDWAREINPUT](https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-hardwareinput)

MOUSEINPUT & KEYBDINPUT
-----------------------
With MOUSEINPUT you simulate mouse-movemements and mouse-clicks. With KEYBDINPUT you simulate Key-down events and key-up events. 
But be careful, you have to know what you do, because with it you can easily screw up your system, and it could behave very weird.
After a key-down of any key, a key-up must follow. e.g.  
Shift-key-down + A-key-down + A-key-up + Shift-key-up: result is the upper case letter "A".
The same applies for mouseinputs after a mouse-down a mouse-up must follow.

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
* collect all data anywhere in the heap memory, and at the time when SendInput is called we copy all data to an array  
* copy all data in and out the array for every editing  
* leave the data in the array and use a pointer instead    

Pointer to a Structure
----------------------
In C we would of course use a pointer to the INPUT structure. This repo Sys_InputSender does it pretty much the same way, it uses the udt-pointer method to create and process the data.

The UDT-Pointer Method for VB
-----------------------------
After playing around with SafeArrays, I figured out how to use it as a pointer to ud-types, and shared my findings with other VB-coders at ActiveVB in 2008.  
How does the udt-pointer method work? We copy the pointer to a [SAFEARRAY](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearray)-descriptor into an empty array of ud-type, and set the pvData-pointer where we want it to point to. 
We create a class for every Input-structure MOUSEINPUT KEYBDINPUT and HARDWAREINPUT. Every class holds a SafeArray-descriptor (TUDTPtr) and a variable of its input-type. No worry, the functions and the type [TUDTPtr](https://github.com/OlimilO1402/Ptr_Pointers/blob/main/Modules/MPtr.bas#L48) we need for this, are already pre-defined in module [MPtr](https://github.com/OlimilO1402/Ptr_Pointers/blob/main/Modules/MPtr.bas) in the repo [Ptr_Pointers](https://github.com/OlimilO1402/Ptr_Pointers)
During creation of the object the pvData of the TUDTPtr points to the internal input-variable. When we add the object to the Array-List (WndInputs), pvData will be set to point to the variable in the array and the data will be copied to the array, with just a private udtype-assignment inside the class. It is even not necessary to get access to the Array in the List itself. The object now gets freed, as we no longer need it. If we edit the data later, a new object will be created. I decided to do it this way, we could of course also leave the object alive and collect it in a second array or collection, just in case we need relations to other objects.

![InputSender Image](Resources/InputSender.png "InputSender Image")
