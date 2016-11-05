# WinErr

### Introduction
WinErr is a simple program that enumerates Windows System Error Codes with their description and allows searching of error descriptions based on the error code.

Also has an export feature that would allow exporting all the enumerated error codes and descriptions to a text file.

There is a precompiled binary available in case you want to test this out immediately.

### Why?
I wrote this program to help me figure out error codes when I work with Windows APIs or programs that do not provide descriptive error descriptions when obtained from ***GetLastError***.

System Error Codes are returned by the ***GetLastError*** function when many functions fail. To retrieve the description text for the error in this program uses the ***FormatMessage*** function with the *FORMAT_MESSAGE_FROM_SYSTEM* or *FORMAT_MESSAGE_IGNORE_INSERTS* flag.

### Requirements
This was written in VB6 a long time ago but can easily be ported to modern .NET.

Runs on Windows XP to Windows 10. Obviously additional error codes have been added to Windows over the years. So you will not find some error codes on older operating systems.

- Open the WinErr.vbp file in VB6 (if you are still using that :D)
- Under File > Make WinErr.exe

### Further Reading
System Error Codes: 
[https://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx](https://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx)