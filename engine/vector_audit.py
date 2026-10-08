"""Non-mutating planar vector diagnostics; curved intersections are candidates."""
import math
from collections import defaultdict
from contour_dxf import distance_to_chord, requires_closed


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


def audit(contours, across_layers=False, limit=1000):
    issues, seen, spans = [],set(),[]
    def add(kind,point,ids,layer,approx=False):
        key=(kind,tuple(sorted(set(ids))),round(point[0],2),round(point[1],2))
        if key not in seen and len(issues)<limit:
            seen.add(key)
            issues.append(dict(kind=kind,point=point,ids=sorted(set(ids)),layer=layer,approx=approx))
    for ci,c in enumerate(contours):
        segs=c['segments']
        if not segs: continue
        if not c['closed'] and requires_closed(c['layer']):
            for point in (segs[0][1][0],segs[-1][1][-1]):
                add('open',point,[c['id']],c['layer'])
        for si,(kind,p) in enumerate(segs):
            if all(math.dist(p[0],x)<1e-9 for x in p):
                add('zero',p[0],[c['id']],c['layer']);continue
            vertices=[p[0],p[-1]] if kind=='L' else flatten(p)
            for j,(a,b) in enumerate(zip(vertices,vertices[1:])):
                if math.dist(a,b)<1e-10:continue
                spans.append((min(a[0],b[0]),max(a[0],b[0]),min(a[1],b[1]),max(a[1],b[1]),a,b,ci,si,j,kind=='B'))
    spans.sort(key=lambda s:s[0])
    active=[]; comparisons=0; complete=True
    for s in spans:
        active=[t for t in active if t[1]+1e-8>=s[0]]
        for t in active:
            if s[2]>t[3]+1e-8 or t[2]>s[3]+1e-8:continue
            ca,cb=contours[s[6]],contours[t[6]]
            if not across_layers and ca['layer']!=cb['layer']:continue
            comparisons+=1
            if comparisons>3_000_000:
                complete=False;break
            found=intersection(s[4],s[5],t[4],t[5])
            if not found:continue
            kind,point=found
            if kind=='intersection' and s[6]==t[6]:
                n=len(ca['segments'])
                neighbors=(s[7]==t[7] and abs(s[8]-t[8])<=1) or abs(s[7]-t[7])==1 or (ca['closed'] and {s[7],t[7]}=={0,n-1})
                # Only a normal shared endpoint is excluded, not a backtracking overlap.
                if neighbors and any(math.dist(point,x)<1e-7 for x in s[4:6]) and any(math.dist(point,x)<1e-7 for x in t[4:6]):continue
            add(kind,point,[ca['id'],cb['id']],ca['layer'],s[9] or t[9])
        if not complete:break
        active.append(s)
    return {'issues':issues,'complete':complete and len(issues)<limit,'segments':len(spans),'comparisons':comparisons}


def ambiguous_endpoints(contours, tolerance):
    points=[(ci,p) for ci,c in enumerate(contours) if not c['closed'] for p in (c['segments'][0][1][0],c['segments'][-1][1][-1])]
    for i,(ci,p) in enumerate(points):
        if sum(i!=j and math.dist(p,q)<=tolerance for j,(_,q) in enumerate(points))>1:
            return True
    return False
