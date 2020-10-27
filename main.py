# -*- coding: utf-8 -*-
import turtle
from shapely.geometry import Polygon

_dict = {}
coord_file = open("lote8.csv","r")

for line in coord_file:
    a = line.splitlines()
    for i in a:
        b = i.split(',')
        _dict[b[0]] = [float(b[1]), float(b[2])]

coord_file.close()

poly = []
for k, v in _dict.items():
    poly.append(tuple(v))

polygon = Polygon(poly)
a = polygon.area
b = polygon.length
print(a,', ',b)

# squirtle = turtle.Turtle()
# squirtle.penup()
# squirtle.setpos(poly2[0])
# squirtle.pendown()
# x, y = poly2[0]
# print(x, y)
# for ponto in poly2:
    # dx, dy = ponto
    # print(ponto)
    # squirtle.goto(x + dx, y + dy)
    # print(x + dx, y + dy)
    # squirtle.goto(x, y)
# turtle.done()