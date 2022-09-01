from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook
from PIL import Image,ImageDraw,ImageFont
import webcolors
import os
import openpyxl

#image_path = "crystal.png"  #need 32-bit depth to get proper rgba values

long_colors = []
colors = []

def img_to_grid(img, dim_x, dim_y):
    grid = []
    palette = []
    with Image.open(img) as image:
        w, h = image.size  # get image dimensions
        pixels = image.load()
        row_w = int(w / dim_x)  # define dimensions of "stitch" based on image dimensions
        row_h = int(h / dim_y)
        n = 0
        for x in range(dim_x):
            row = []
            for y in range(dim_y):
                box = (y * row_w, x * row_h, y * row_w + row_w, x * row_h + row_h)  # define box (range of pixels) to crop out of original image
                output = image.crop(box)  # crop
                row.append(get_avg_color(output))
                
                if get_avg_color(output) not in palette:
                    palette.append(get_avg_color(output))
                else:
                    continue
                # print(palette)
                # name, ext = os.path.splitext(img)
                # output_path = f"output\\{name}_{str(n)}{ext}"
                # output.save(output_path)  # save to new directory for later
                n+=1
            
            grid.append(row)
            
    return grid, palette


def grid_to_sheet(grid, long_colors):
    sheet = "test.xlsx"
    wb = Workbook()
    ws = wb.active
    border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    
    for x in range(len(grid)):
        for y in range(len(grid[x])):
            ws.column_dimensions[f'{get_column_letter(y+1)}'].width = 3
            cell = ws[f'{get_column_letter(y+1)}{x+1}']
            cell.fill = PatternFill("solid", fgColor=f'{grid[x][y].strip("#")}')
            cell.border = border
    
    wb.save(sheet)
    
    colors = list(set(long_colors))
    font = ImageFont.load_default().font
    palette = Image.new(mode="RGB", size=(1000,120))

    for i in range(len(colors)):
        newt = Image.new(mode="RGB", size=(1000//len(colors),100), color=colors[i])
        draw = ImageDraw.Draw(newt)
        draw.text((5,5), str(colors[i]), fill="white", font=font)
        palette.paste(newt, (i*1000//len(colors),10))

    palette.save("palette.png")
    

def get_avg_color(img):  # scan the image and return the most common color in hex
    long_colors = []
    w, h = img.size
    pixels = img.load()
    for y in range(h):
        for x in range(w):
            try:
                #print(pixels[x,y])
                r, g, b, a = pixels[x,y]
                if a == 0:
                    long_colors.append(webcolors.rgb_to_hex((255,255,255))) #white for blank pixels
                else:
                    long_colors.append(webcolors.rgb_to_hex((r, g, b)))
            except TypeError:
                print(pixels[x,y])
    
    most_common = max(set(long_colors), key = long_colors.count)
    return most_common
    

def get_palette(img):
    with Image.open(img) as image:
        w, h = image.size
        pixels = image.load()
        for y in range(h):
            for x in range(w):
                try:
                    #print(pixels[x,y])
                    r, g, b, a = pixels[x,y]
                    long_colors.append(webcolors.rgb_to_hex((r, g, b)))
                except TypeError:
                    print(pixels[x,y])

    colors = list(set(long_colors))
    font = ImageFont.load_default().font
    palette = Image.new(mode="RGB", size=(1000,120))

    for i in range(len(colors)):
        newt = Image.new(mode="RGB", size=(1000//len(colors),100), color=colors[i])
        draw = ImageDraw.Draw(newt)
        draw.text((5,5), str(colors[i]), fill="white", font=font)
        palette.paste(newt, (i*1000//len(colors),10))

    palette.save("palette.png")

if __name__ == "__main__":
    img_path = input("Enter path to image file: ")
    dimension_x, dimension_y = input("Enter stitch dimensions (ex: 20x20)\nInput: ").split("x")
    print("Working...")
    grid, palette = img_to_grid(img_path, int(dimension_x), int(dimension_y))
    grid_to_sheet(grid, palette)
    print("DONE!")
    
    
