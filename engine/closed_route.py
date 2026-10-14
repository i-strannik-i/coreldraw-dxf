"""Exact shortest connections for a small set of reversible open contours.

All original segments are traversed once. Added straight segments produce one
closed contour. This is cutting geometry, not a sequence of rapid moves.
"""
import math


def oriented_segments(contour, reverse=False):
    segments=contour['segments']
    if not reverse:return [(kind,list(points)) for kind,points in segments]
    return [(kind,list(reversed(points))) for kind,points in reversed(segments)]


def shortest_closed_route(contours):
    count=len(contours)
    if not 1<=count<=12:
        raise ValueError('Точный тестовый расчёт поддерживает от 1 до 12 открытых контуров.')
    if any(c['closed'] or not c['segments'] for c in contours):
        raise ValueError('Нужны непустые открытые контуры.')
    endpoints=[(c['segments'][0][1][0],c['segments'][-1][1][-1]) for c in contours]
    if any(not all(math.isfinite(v) for v in point) for ends in endpoints for point in ends):
        raise ValueError('Некорректные координаты.')
    # A closed route has no distinguished first contour. Reversing the entire
    # cycle also preserves length, so contour 0 can be fixed in forward order.
    start=(1,0,0)
    costs={start:0.0};parents={}
    for mask in range(1,1<<count,2):
        for last in range(count):
            for direction in (0,1):
                state=(mask,last,direction)
                if state not in costs:continue
                end=endpoints[last][1-direction]
                for following in range(1,count):
                    if mask & (1<<following):continue
                    for reverse in (0,1):
                        target=(mask|(1<<following),following,reverse)
                        distance=costs[state]+math.dist(end,endpoints[following][reverse])
                        if distance<costs.get(target,math.inf):
                            costs[target]=distance;parents[target]=state
    full=(1<<count)-1
    final=min((s for s in costs if s[0]==full),key=lambda s:
              costs[s]+math.dist(endpoints[s[1]][1-s[2]],endpoints[0][0]))
    order=[];state=final
    while True:
        order.append((state[1],bool(state[2])))
        if state==start:break
        state=parents[state]
    order.reverse()
    segments=[];bridges=[]
    for position,(index,reverse) in enumerate(order):
        current=oriented_segments(contours[index],reverse)
        segments.extend(current)
        next_index,next_reverse=order[(position+1)%count]
        a=current[-1][1][-1];b=endpoints[next_index][int(next_reverse)]
        if math.dist(a,b)>1e-9:
            segments.append(('L',[a,a,b,b]));bridges.append((a,b))
    baseline=sum(math.dist(endpoints[i][1],endpoints[(i+1)%count][0]) for i in range(count))
    return dict(order=order,segments=segments,bridges=bridges,
                added_length=sum(math.dist(a,b) for a,b in bridges),
                baseline_added_length=baseline,exact=True)
