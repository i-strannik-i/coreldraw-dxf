"""Conservative full-object duplicates, independent of traversal direction."""
from collections import defaultdict

DUPLICATE_PREFIX = '_DXF_DUPLICATES'


def minimum_rotation(items):
    n=len(items)
    if not n:return items
    doubled=items+items
    i,j,k=0,1,0
    while i<n and j<n and k<n:
        a,b=doubled[i+k],doubled[j+k]
        if a==b:k+=1;continue
        if a>b:
            i+=k+1
            if i==j:i+=1
        else:
            j+=k+1
            if i==j:j+=1
        k=0
    start=min(i,j)
    return doubled[start:start+n]


def contour_key(contour):
    # Numerical noise only, NOT the user's machining/join tolerance.
    forward=tuple((kind,tuple(tuple(round(v,9) for v in p) for p in points))
                  for kind,points in contour['segments'])
    backward=tuple((kind,tuple(reversed(points))) for kind,points in reversed(forward))
    if contour['closed']:
        forward=minimum_rotation(forward);backward=minimum_rotation(backward)
    return bool(contour['closed']),min(forward,backward)


def duplicate_groups(contours):
    objects=defaultdict(list)
    for contour in contours:
        objects[(contour['layer'],contour['id'])].append(contour)
    geometry=defaultdict(list)
    for (layer,sid),parts in objects.items():
        if not parts or any(not p['segments'] for p in parts):continue
        key=(layer,tuple(sorted(contour_key(p) for p in parts)))
        geometry[key].append(sid)
    groups=[]
    for (layer,_),ids in geometry.items():
        if len(ids)<2:continue
        ids=sorted(ids)
        groups.append(dict(kind='duplicate',layer=layer,ids=ids,keep=ids[0],copies=ids[1:],
                           point=objects[(layer,ids[0])][0]['segments'][0][1][0],approx=False))
    return groups
