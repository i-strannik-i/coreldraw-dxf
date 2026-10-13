"""Non-mutating planar vector diagnostics; curved intersections are candidates."""
import math
from collections import defaultdict
from contour_dxf import distance_to_chord, requires_closed
from duplicate_geometry import duplicate_groups, duplicate_removals

ISSUE_ORDER = ('duplicate', 'zero', 'overlap', 'intersection', 'open')


def flatten(points, error=.005, depth=0):
    if max(distance_to_chord(p, points[0], points[-1]) for p in points[1:3]) <= error:
        return [points[0], points[-1]]
    if depth >= 24:
        raise ValueError('Curve too complex for bounded validation')
    levels = [list(points)]
    while len(levels[-1]) > 1:
        levels.append([((a[0]+b[0])/2, (a[1]+b[1])/2) for a,b in zip(levels[-1],levels[-1][1:])])
    left, right = [p[0] for p in levels], [p[-1] for p in reversed(levels)]
    return flatten(left,error,depth+1)[:-1]+flatten(right,error,depth+1)


def cross(a,b):
    return a[0]*b[1]-a[1]*b[0]


def intersection(a,b,c,d):
    r=(b[0]-a[0],b[1]-a[1]); s=(d[0]-c[0],d[1]-c[1]); q=(c[0]-a[0],c[1]-a[1])
    den=cross(r,s)
    if abs(den) > 1e-12:
        t,u=cross(q,s)/den,cross(q,r)/den
        if -1e-9 <= t <= 1+1e-9 and -1e-9 <= u <= 1+1e-9:
            return 'intersection',(a[0]+t*r[0],a[1]+t*r[1])
        return None
    length2=r[0]**2+r[1]**2
    if length2 < 1e-18 or abs(cross(q,r)) > 1e-8*math.sqrt(length2):
        return None
    ts=[((p[0]-a[0])*r[0]+(p[1]-a[1])*r[1])/length2 for p in (c,d)]
    lo,hi=max(0,min(ts)),min(1,max(ts))
    if (hi-lo)*math.sqrt(length2) > 1e-7:
        t=(lo+hi)/2
        return 'overlap',(a[0]+t*r[0],a[1]+t*r[1])
    return None


def audit(contours, *, tolerance=.1, limit=1000, progress=None, cancelled=None, layer_done=None):
    if not math.isfinite(tolerance) or not .001 <= tolerance <= 1:
        raise ValueError('Допуск проверки: от 0,001 до 1 мм.')
    layers = defaultdict(list)
    for contour in contours:
        layers[contour['layer']].append(contour)
    result = dict(issues=[], complete=True, segments=0, comparisons=0, layers=[])
    for index, (layer, items) in enumerate(layers.items()):
        def report(stage, done, total):
            if cancelled and cancelled():
                raise InterruptedError('Проверка отменена. Итоги завершённых слоёв сохранены в журнале.')
            if progress:
                progress(layer, index + 1, len(layers), stage, done, total)
        # Each layer owns its own budget; one busy layer cannot hide the rest.
        part = _audit_layer(items, tolerance, limit, report)
        part=dict(part,layer=layer,count=len(items))
        result['layers'].append(part)
        if layer_done:layer_done(part)
        result['issues'].extend(part['issues'])
        result['segments'] += part['segments']
        result['comparisons'] += part['comparisons']
        result['complete'] &= part['complete']
    result['issues'].sort(key=lambda issue: ISSUE_ORDER.index(issue['kind']))
    return result


def _audit_layer(contours, tolerance, limit, report):
    issues, seen, spans = [],set(),[]
    report('Поиск дублей',0,len(contours))
    duplicates=duplicate_groups(contours,lambda done,total:report('Поиск дублей',done,total))
    whole_copies,partial_copies=duplicate_removals(duplicates)
    issues.extend(duplicates[:limit])
    def add(kind,point,ids,layer,approx=False):
        key=(kind,tuple(sorted(set(ids))),round(point[0],2),round(point[1],2))
        if key not in seen and len(issues)<limit:
            seen.add(key)
            issues.append(dict(kind=kind,point=point,ids=sorted(set(ids)),layer=layer,approx=approx))
    ordinals=defaultdict(int)
    retained=[]
    for ci,c in enumerate(contours):
        subpath=c.get('subpath',ordinals[c['id']]);ordinals[c['id']]+=1
        if c['id'] in whole_copies or subpath in partial_copies.get(c['id'],set()):continue
        retained.append(c)
        report('Подготовка кривых', ci, len(contours))
        segs=c['segments']
        if not segs: continue
        for si,(kind,p) in enumerate(segs):
            if all(math.dist(p[0],x)<1e-9 for x in p):
                add('zero',p[0],[c['id']],c['layer']);continue
            vertices=[p[0],p[-1]] if kind=='L' else flatten(p, tolerance)
            for j,(a,b) in enumerate(zip(vertices,vertices[1:])):
                if math.dist(a,b)<1e-10:continue
                spans.append((min(a[0],b[0]),max(a[0],b[0]),min(a[1],b[1]),max(a[1],b[1]),a,b,ci,si,j,kind=='B'))
    spans.sort(key=lambda s:s[0])
    active=[]; comparisons=0; complete=True
    for number,s in enumerate(spans):
        if number % 100 == 0:
            report('Поиск пересечений', number, len(spans))
        active=[t for t in active if t[1]+1e-8>=s[0]]
        for t in active:
            if s[2]>t[3]+1e-8 or t[2]>s[3]+1e-8:continue
            ca,cb=contours[s[6]],contours[t[6]]
            if ca['layer']!=cb['layer']:continue
            comparisons+=1
            if comparisons % 10000 == 0:
                report('Поиск пересечений', number, len(spans))
            if comparisons>3_000_000:
                complete=False;break
            found=intersection(s[4],s[5],t[4],t[5])
            if not found:continue
            kind,point=found
            if kind=='intersection':
                def open_end(contour,span):
                    if contour['closed']:return False
                    segments=contour['segments']
                    return ((span[7]==0 and math.dist(point,segments[0][1][0])<1e-7) or
                            (span[7]==len(segments)-1 and math.dist(point,segments[-1][1][-1])<1e-7))
                # A meeting of open ends is a join candidate, not a crossing.
                if open_end(ca,s) and open_end(cb,t):continue
            if kind=='intersection' and s[6]==t[6]:
                n=len(ca['segments'])
                neighbors=(s[7]==t[7] and abs(s[8]-t[8])<=1) or abs(s[7]-t[7])==1 or (ca['closed'] and {s[7],t[7]}=={0,n-1})
                # Only a normal shared endpoint is excluded, not a backtracking overlap.
                if neighbors and any(math.dist(point,x)<1e-7 for x in s[4:6]) and any(math.dist(point,x)<1e-7 for x in t[4:6]):continue
            add(kind,point,[ca['id'],cb['id']],ca['layer'],s[9] or t[9])
        if not complete:break
        active.append(s)
    report('Поиск открытых контуров',0,len(contours))
    for ci,c in enumerate(retained):
        report('Поиск открытых контуров',ci,len(contours))
        segs=c['segments']
        if segs and not c['closed'] and requires_closed(c['layer']):
            for point in (segs[0][1][0],segs[-1][1][-1]):
                add('open',point,[c['id']],c['layer'])
    report('Готово',len(spans),len(spans))
    return {'issues':issues,'complete':complete and len(issues)<limit,'segments':len(spans),'comparisons':comparisons}


def ambiguous_endpoints(contours, tolerance):
    points=[(ci,p) for ci,c in enumerate(contours) if not c['closed'] for p in (c['segments'][0][1][0],c['segments'][-1][1][-1])]
    for i,(ci,p) in enumerate(points):
        if sum(i!=j and math.dist(p,q)<=tolerance for j,(_,q) in enumerate(points))>1:
            return True
    return False
