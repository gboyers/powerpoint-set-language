# powerpoint-set-language
This is some AppleScript to set the language on all objects/shapes/boxes in a Microsoft PowerPoint presentation.

Confirmed to work on:
* Mac OS Cataline 10.15
* Microsoft PowerPoint for Mac 16 (Microsoft 365)

## Why do I need this?

Changing the language on PowerPoint is HARD. 

If you want your presentation to be in UK English, but you're using slides / text boxes created in US English, you'll end up with spelling errors and red lines all over.

You can set your default laniage in PowerPoint... but that only applies to new presentations.

You can change the language of an existing slide... that only applies to new objects created on it. It won't change existing shapes or text boxes.

You can select the outline of the whole presentation and change the language... but that only applies to text that is in content placeholders (title or body), not other shapes or text boxes.

So you need to go through every single slide individually, select all the shapes & text boxes on each, and change the language on all of them. 

This AppleScript automates that for you.

## What does this script do?

* Opens Powerpoint
* Goes to Slide 1
* Selects all shapes
* Opens the Language Window
* Chooses the language from the list that matches the code on line 28
* Moves onto Slide 2
* Repeats until the last slide

## How do I use it?

1. Download the script
    * Click Code > Download Zip & extract the zip
2. Open the PowerPoint slideshow you wish to change the 
3. Open `powerpoint-set-language.scpt`
    * This should open in your Mac's Script Editor
5. If required, change the language on line 28
    * Choose a language exactly as listed in the PowerPoint language window, such as: `English (United States)` or `Chinese (China)`
    * You can see a full list here - https://www.codetwo.com/admins-blog/list-of-office-365-language-id/
6. Hit Run ▶️ 
7. If you get a message saying you need to grant Script Editor accessibility access
  1. Make sure you have Administrator rights
  2. Open System Preferences > Security & Privacy > Privacy
  3. Click the padlock (if required) to unlock this screen
  4. Go to Accessibility, and tick 'Script Editor'

If you need to stop it running, you can click Cancel on the Language window, or go back to Script Editor and hit stop ⏹ 


## If it doesn't work?

* Make sure you're on a Mac. If you're on Windows, you'll need to find a VBA script that does the same thing, such as: https://gist.github.com/erikvullings/f197ee68c6119a28d070926175690812
* So far this has only been tested on Mac OS Catalina, and PowerPoint 365 v16. It may work differently on other versions.
* Make sure you have given Accessibility access to Script Editor (step 7 above)
* If you can't grant Accessibility access, try clicking the padlock to unlock the screen. If you can't do this, speak to your system administrator.
* Make sure that you have PowerPoint open, with the document you want
* Make sure you have no other PowerPoint windows or dialog boxes open
* If you've changed the language, make sure it exactly matches the text shown in the PowerPoint list - and make sure you haven't accidentally removed/replaced the quotes in the script

## But I still have a problem

You can raise an issue on GitHub here: https://github.com/gboyers/powerpoint-set-language/issues 

Share the details of the error you get, and screenshots if you can. 

No promises, but someone might be able to help you out. 
