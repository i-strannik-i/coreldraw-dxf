"""Order vector contours without changing their geometry or direction."""
import math


def route_length(contours, order):
    return sum(math.dist(contours[a]['segments'][-1][1][-1],
                         contours[b]['segments'][0][1][0])
               for a,b in zip(order,order[1:]))


def optimize_route(contours):
    """Greedy travel reduction, with conservative inside-before-outside ordering.

    Control-point bounds enclose Beziers. A closed contour enclosing another
    contour's bounds is scheduled after it. Starts and directions stay intact.
    Distances exclude approach from the machine origin and the final return.
    """
    count=len(contours)
    original=list(range(count))
    if count<2:return original,0.0,0.0
    starts=[c['segments'][0][1][0] for c in contours]
    ends=[c['segments'][-1][1][-1] for c in contours]
    bounds=[]
    for c in contours:
        points=[p for _,segment in c['segments'] for p in segment]
        xs,ys=zip(*points)
        bounds.append((min(xs),min(ys),max(xs),max(ys)))
    dependencies=[set() for _ in contours]
    for i,outer in enumerate(contours):
        if not outer['closed']:continue
        a=bounds[i]
        for j,b in enumerate(bounds):
            if i!=j and a!=b and a[0]<=b[0] and a[1]<=b[1] and a[2]>=b[2] and a[3]>=b[3]:
                dependencies[i].add(j)
    def build(nearest):
        remaining=set(original);done=set();order=[]
        while remaining:
            available=[i for i in remaining if dependencies[i]<=done]
            if nearest and order:
                chosen=min(available,key=lambda i:(math.dist(ends[order[-1]],starts[i]),i))
            else:chosen=min(available)
            order.append(chosen);done.add(chosen);remaining.remove(chosen)
        return order
    baseline=build(False)
    greedy=build(True)
    best=min((baseline,greedy),key=lambda order:route_length(contours,order))
    return best,route_length(contours,original),route_length(contours,best)


def join_summary(before,after,tolerance,route=None):
    joined=sum(before)-sum(after)
    closed=max(0,after[0]-before[0])
    lines=[f'Соединений: {joined}. Замыканий: {closed}.',
           f'Результат: замкнутых — {after[0]}, открытых — {after[1]}.']
    if after[1]:
        lines.append(f'Оставшиеся концы не соединяются в допуске {tolerance:g} мм.')
    if route is not None:
        old,new=route
        factor,unit=(1000,'м') if max(old,new)>=1000 else (1,'мм')
        distance=lambda value:f'{value/factor:.1f}'.replace('.',',')
        lines.append(f'Расчётный путь: {distance(old)} → {distance(new)} {unit}.')
        if abs(old-new)<1e-6:lines.append('Сокращения пути нет.')
        elif old>0:
            percent=abs((old-new)/old*100)
            lines.append(f'{"Сокращение" if new<old else "Увеличение"}: {percent:.0f}%.')
        else:lines.append('Порядок изменён с учётом вложенных контуров.')
    return '\n'.join(lines)
