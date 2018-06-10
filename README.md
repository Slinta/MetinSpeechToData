﻿# *Metin2* Drop Logger
---
This is a simple program for item logging in **Metin2.**

## How to use it:
* [Download it](https://github.com/Slinta/MetinSpeechToData/releases "Releases")
* Have a microphone
* Have a way to open .xlsx files >> supported by us **Microsoft Excel** or **LibreOffice Calc** (**OpenOffice** is not supported, but may work)
* Have a somewhat recent version of **.NET Framework**

There are two executables in the folder you download, one is for the recognition and the other for merging session files into main one.
If you do not need/want to have total sums and averages across sessions then you will not need to use it.

**After launching it for the first time, it is recommended to close it and edit the configuration file that appeared in the main folder**!\
It contains a few interesting settings and mainly you define what program you intend to use to open .xlsx files, because they are not fully compatible..

You probably noticed some **.definition** files, they are used to store recognizer grammar, the syntax for new item/enemy definition is written at the top.

While the program is configured and running, you can say any of the items you defined, and they will be put in a session file.\
If you do not want to use voice to control this program, it is possible to use hotkeys instead,\
hotkey definition is now possible only via a text editor,
but I plan to implement a definition from inside the program.

When you decide to quit the recognition, you will be prompted to type a name of this session, after that you will find it in *"Sessions"* folder in current directory.


## Currently We are working on:
* Fixing typos in the program
* Localization ?
* A way to know the amount of money you started with
								someone who will produce quality .definition files ;)

If you want to contribute, feel free to create a pull request.

If you know how Deep learning NNs work, then branch NeuralNetwork_Enemy_Recognizer is yours, it is converging to 0.5 for simple XOR with 2in 2hidden 1out.
I have no idea why...

# Contributions:
1. Michal-MK - main developer
2. Slinta - developer/tester
3. "Possibly your name here"