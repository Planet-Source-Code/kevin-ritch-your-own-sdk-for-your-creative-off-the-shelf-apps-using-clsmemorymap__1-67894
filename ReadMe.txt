Hi folks,

Provide your own SDK for your creative off-the-shelf apps using clsMemoryMap!

With clsMemoryMap (Class Module employing RtlMoveMemory et al) you can SHARE memory between 
your application and others. It is therefore, a darned neat way to provide interactive 
communications between your retail application and Add-Ons that can be developed by 3rd parties 
for your package!

Of course, this is nothing new (I have no idea who knocked out the clever Class Module) but 
what I am showing you all here, is a rudementary method to provide a "Software Development Kit" 
for your cool applications that can be remotely controlled in a fashion similar to DDE. Sans 
the issues you may otherwise encounter with DDE.

This submission comprises two very simple VB6 programs.

The FIRST VB6 App is called "WidgetView" a useless thing that has a simple list of 5 names. 
WidgetView has a timer called MySDK that examines a SHARED MEMORY (vb6Request) and updates a 
Label. When the label changes, it invokes a request from the listbox and submits the data back 
to the AddOn program, just by "poking" the result into a SHARED MEMORY (vb6Result).

The SECOND VB6 App is called "MyAddOn". It does very little. The user picks a number between 1 
and 5 and sets the SHARED MEMORY (VB6Request) with that value, and uses a similar timer called 
MySDK, to examine changes to the Vb6Result memory variable.

It is, as I say, rudimentary, but as you make up your own rules and control them with 
CASE/SELECT statements, you can build up a pretty groovy SDK for your App in next to no time!

Enjoy,
Kevin Ritch
V8Software.com
631-961-0594