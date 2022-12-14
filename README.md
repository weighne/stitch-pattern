# stitch-pattern
Turn an image into a cross stitch pattern using python

## Dependencies

* [Pillow](https://pypi.org/project/Pillow/)
* [webcolors](https://pypi.org/project/webcolors/)
* [openpyxl](https://pypi.org/project/openpyxl/)

## How to use this thing

1. Launch the script
2. Type in the path to the image (if it's in the same directory as the script, just use the file name)
3. Entire the desired stitch pattern dimensions
4. Wait a few seconds for the script to work
5. Once complete, the script should spit out an excel spreadsheet and a PNG with the color palette used for the pattern

Depending on the image you use, the palette might look a bit busy. I just wanted it present in case you need to edit the final sheet :). 

## Q&A

### Why a spreadsheet?

I wanted the final pattern to be easily modifiable and a spreadsheet makes it easy to just pop in and change out a few cells

### Why did you do this?

I like cross-stitch, but I don't like having to dig around for cool patterns or spend time making my own patterns, so this automates some of the busy work. 

### This thing is a mess!

You're correct, I will be cleaning it up as I go (and probably adding more features and perhaps a GUI).
