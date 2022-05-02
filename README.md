# EX3-sampler-automation
windows desktop application (C#) for adding functionalities to the EmulatorX 3 (EMU) software sampler (via gui automation)

I wrote this stuff the XP way so it really doesn't follow good practices too much :)
feel free not to use it because it really is for my own purpose only !

## functionalities
- allow full Voice Processing data structure copy and paste operation
- allow to convert a "link based preset" to a "voice and zone based preset"
- as the application gives the programmer the ability to read and write the EX3 data structures, new functionalities cound be developped easily
- 
## usage
Once started, The application waits for the EmulatorX3 application to be locally opened by the user
is present, the application will then show it's interface where the user can access the added features

## caution
manipulating data thru the GUI is not the ideal way but on the case ofan external software like EX3 it can help improve some time consuming tasks. It also can stuckj your machine in endless loops and seriously damage your system so if not sure of what you are doing with this piece of code: simply don't run it!

