# EZ-RDP
Simple VBA scripts for managing a multitude of RDP clients.

## What does it do?
- Spawns the RDP client on your desktop and autofills the name given from "Machine Name".
- Searches the list based on Room # and Building, then returns the concatenated PC name and domain.
- Prints your search result in a message box and in the empty "Machine Name" box.

## What does it not do?
- Handle search errors very well. If you do not enter a value for Room #, or enter something other than an integer, you will get a type-mismatch error.

## Other Considerations
The image file is an example of how I have my data set up. It has been blurred for the sake of privacy - but all you need is the headers to recreate it with your data.

![alt text](https://github.com/eggsnbacon97/EZ-RDP/blob/master/ez%20rdp%20example.png)
