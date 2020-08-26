import xlrd                # for reading data from Excel file
from math import sqrt      
from turtle import Turtle, Screen   # for graphics design 
import time
import turtle

c=11                                # 10 cells


def fun(j,k,mylist):
    file_location="C:/Users/HP/Desktop/DRS/Book"
    file_location=file_location+str(k);
    file_location=file_location+".xlsx";
    workbook=xlrd.open_workbook(file_location)    # Reading data from excel file
    sheet=workbook.sheet_by_index(0)
    r=sheet.nrows
    c=sheet.ncols
   
    temp=0
    dict={}
    for i in range(1,r):
        dict[sheet.cell_value(i,j)]=sheet.cell_value(i,0)

    for user,item in dict.items():
        if(user>temp):
            temp=user
            fina=item

    mylist.append(fina)
       
       


def main():                         # Main function
    for k in range(1,25):            # Picking the time
        mylist=[]                   # lisr declaration
        for i in range(1,c):        # Reading the dataset every hour
            fun(i,k,mylist)
            
        print(mylist)

    

        ROOT3_OVER_2 = sqrt(3)/2   # Graphics Representation
        FONT_SIZE = 12
        FONT = ('Arial', FONT_SIZE, 'normal')
        

        SIDE = 90  # the scale used for drawing

        # Convert hex coordinates to rectangular
        def hex_to_rect(coord):
            v, u, w = coord
            x = -u / 2 + v - w / 2
            y = (u - w) * ROOT3_OVER_2
            return x * SIDE, y * SIDE

        def hexagon(turtle, radius, color, label):
            clone = turtle.clone()  # so we don't affect turtle's state
            xpos, ypos = clone.position()
            clone.setposition(xpos - radius / 2, ypos - ROOT3_OVER_2 * radius)
            clone.setheading(-30)
            clone.color('black', color)
            clone.pendown()
            clone.begin_fill()
            clone.circle(radius, steps=6)
            clone.end_fill()
            clone.penup()
            clone.setposition(xpos, ypos - FONT_SIZE / 2)
            clone.write(label, align="center", font=FONT)

        # Initialize the turtle
        tortoise = Turtle(visible=False)
        tortoise.speed('fastest')
        tortoise.penup()
        Dict2 = {'Live linguistic news': 'Red', 'Live sports event': 'Grey', 'Cinematic movies': 'Blue','Anime/Cartoon': 'Pink', 'Study related material': 'Orange', 'Company related material': 'Cyan', 'Web Series': 'Yellow', 'Fitness videos': 'Green','Online gaming': 'White'} 
        coords = [[0, 0, 0], [0, 1, -1], [-1, 1, 0], [-1, 0, 1], [0, -1, 1], [1, -1, 0], [1, 0, -1],[-2, 1, 1],[4,4,1],[1,1,4]]
        labels=mylist
        #colors = [Dict2[8], Dict2[3], Dict2[8], Dict2[8], Dict2[1], Dict2[8], Dict2[8],Dict2[1],Dict2[1],Dict2[1]]
        colors = [Dict2[mylist[0]], Dict2[mylist[1]], Dict2[mylist[2]], Dict2[mylist[3]], Dict2[mylist[4]], Dict2[mylist[5]], Dict2[mylist[6]], Dict2[mylist[7]], Dict2[mylist[8]], Dict2[mylist[9]]]

        # Plot the points
        for hexcoord, color, label in zip(coords, colors, labels):
            tortoise.goto(hex_to_rect(hexcoord))
            hexagon(tortoise, SIDE, color, label)

        # Wait for the user to close the window
        screen = Screen()
        time.sleep(5)
    screen.exitonclick()
        

            


if __name__=="__main__":
    main()
