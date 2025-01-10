# Uphide
Back when Windows 10 was first released and for several years after, Microsoft provided a tool allowing you to show and hide updates if there were updates that were causing problems, called wushowhide. This tool is hard to find and the platform it is built on is now deprecated, so I wrote this simple Windows program to hide any unwanted Windows updates. Simply run it, wait for it to list all your available updates, and type the numbers of the updates to hide, seperated by spaces.

## Note
Due to how the Windows update API works, this program requires administrative privilages to run. A .res file, called uac.res, is included in the repository for this purpose. This has to be linked to have the program automatically ask for administrative rights.

## Download
[Download Uphide](https://quinbox.xyz/files/uphide.exe)

## Todo
For some reason, after hiding at least one update, if you have error reporting set to on, Windows will show a "this program has stopped working error". This doesn't really effect anything and the updates still hide just fine, but I would like to fix it and at the time of writing it is being actively investigated.
