import PIL
from PIL import Image, ImageFont, ImageDraw  
import xlrd 
  
# Give the location of the file 
loc = ("mydata.xlsx") 
  
# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  
# For row 0 and column 0 
print(sheet.nrows)
print(sheet.ncols)
for i in range(sheet.nrows):
    Name = (sheet.cell_value(i, 2))
    Name = Name.capitalize()
    Name = Name.center(28, " ")
    #print(sheet.cell_value(i, 0))
    # creating a image object 
    #print(type(Name))
    #print(len(Name))

    #mod = len(Name) % 2
    #if mod > 0:
    #    count = 25
    #else: 
    #    count = 26
        
    #while True:
    #    if len(Name) < count:
    #        Name = " {} ".format(Name)
    #    else:
    #        break
    
    image = Image.open(r'template.jpg')  
      
    draw = ImageDraw.Draw(image)  
      
    # specified font size 
    font = ImageFont.truetype(r'MoongladeDemoBold-jOzM.ttf', 140)  
      
    text = Name
      
    # drawing text size 
    draw.text((780, 1200), text, fill ="Black", font = font, align ="center")  
      
    image.save("Participants\{}.jpg".format(Name))  