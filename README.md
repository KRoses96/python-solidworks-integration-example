# Python - Solidworks Integration Example

This repo is an example of how you much you can really do once you connect your solidworks with python and how much you can really push it to, for more information on the actual method and the way I'm connecting solidworks and python check the following repo: 

[GitHub - KRoses96/python-solidworks-integration](https://github.com/KRoses96/python-solidworks-integration)

This started as a simple macro that simply turned a solidworks BOM into an excel but ended growing astronomically, my only recommendation for anyone starting a project like this or simply a macro is that you should think about scalability of it, it can get ugly really fast.

---

## Disclaimer

All solidworks macros are as txts since github does not recognize swp.

This is not a github repo for you to clone and use, this is mainly a showcase of how much you can really do just from using solidworks and python together, I'll pin point all the biggest features and the script that handles that, if it's something you intend to implement maybe look at how I did it.

---

### Tech Stack

- Python

- VBA/Solidworks API

- Tkinter/PyQt (used it once to try it out)

- Pandas 

---

### Scripts

Most of the VBA macros are very simple interactions with the model that consist of extraction of data + running python scripts/executables, the more complicated ones are edited macros from this library:

[Library of macros and scripts to automate SOLIDWORKS](https://www.codestack.net/solidworks-tools/)

- macrorun.py:
  
  - what connects our python scripts with solidworks, for more information on this setup, [GitHub - KRoses96/python-solidworks-integration](https://github.com/KRoses96/python-solidworks-integration)

- cutlist.py:
  
  - Creates a list and organizes/analyses the data for all the weld profiles used in the assembly

- sheet_metal.py:
  
  - Analyses all dxfs and attributes them different types of machines
  
  - Creates list of all DXF
  
  - Creates a list of the amount of sheet metal needed for the project and their sizes
  
  - Creates an image to review the DXFs files
  
  - Creates structured folders for every DXF file to ease the cutting of them

- mat.py:
  
  - Creates a list of every item from an assembly that has the property "Paraf"

- data_pass.py:
  
  - Transfers all data inside a software made folder to the production folder, to then be read by their interface for further investigation, for how I was reading the data check the viewer folder, it's a simple tkinter GUI to help with the reviewing process of all this information.
- print.py:
  - Turns dwg files into PDFs automatically, merges pdfs for easier printing and creates a list of all drawings
---

### What next?

I'm working on a web app for rendering and modifying CAD files with a very similar approach. Future plans include releasing some tools as open-source, so if you're interested, send me a message.
