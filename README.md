![MS Office Macro Bleach](docs\images\bleach.jpg)


# MS Office Macro Bleach

A command-line tool designed to detect and purge any and all macros and dynamic content from MS Office files.

It currently supports '.docm' and'.docx' files and will continue to expand its file support with every released version. The next file types to be supported will be '.pdf'' and '.doc'.


## Problem

VBA and OLE content in MS Office files can, and have sometimes been made to, act as vehicles for malware delivery.

Microsoft has previously attempted to protect users from macros by disabling them by default.  However, anybody is able to enable macros in an MS Office file before sending them on to a potential victim.

This Python tool aims to detect and remove any of this potentially malicious content from given files.


## Solution

A command-line program written in modern Python (3.10+) that is capable of locating and removing macros and dynamic content from a variety of files.

It should support all the common Office Open XML formats (e.g. '.pptx', .docx', '.xlsx', etc.) as well as the 'legacy' MS binary file formats, like '.doc' and '.ppt'.
