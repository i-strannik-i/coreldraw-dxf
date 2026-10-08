"""Collision-free number boxes; markers themselves retain exact coordinates."""
SIDES={'zero':'top','overlap':'left','intersection':'right','open':'bottom','duplicate':'top-right'}


def overlaps(a,b,gap=.5):
    return a[0]<b[2]+gap and b[0]<a[2]+gap and a[1]<b[3]+gap and b[1]<a[3]+gap


def place_label(point,kind,width,height,occupied,markers):
    x,y=point;side=SIDES[kind]
    step=max(width,height)+1
    distance=2
    while True:
        if side=='left':left,bottom=x-distance-width,y-height/2
        elif side=='right':left,bottom=x+distance,y-height/2
        elif side=='top':left,bottom=x-width/2,y+distance
        elif side=='bottom':left,bottom=x-width/2,y-distance-height
        else:left,bottom=x+distance,y+distance
        box=(left,bottom,left+width,bottom+height)
        if not any(overlaps(box,other) for other in occupied) and not any(overlaps(box,m,0) for m in markers):
            return box
        distance+=step
