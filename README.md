# SAN FMS

Commerce IT is in the process of setting up an web-based Storage Area Network (SAN) for the use of both staff and students. SAN FMS (Storage Area Network Folder Management System) is an application written specifically for this project and is designed to run on the webserver from which the SAN operates, running its operation every 12 hours in which it manages the creation and removal of both student and staff data folders.

SAN FMS basically sends a request to a specially set up webpage service that returns both the staff numbers and student numbers that all belong to the Commerce faculty. Parsing these results, SAN FMS then creates or removes folders based on whether or not a name is present in the parsed list.

The operation is designed to run twice a day, but you can actually specify it to run more often should you so wish.

Created by Craig Lotter, November 2007

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2005
Implements concepts such as Webpage Downloading, Text File Parsing, Folder Manipulation and Scheduling.
Level of Complexity: Very Simple
