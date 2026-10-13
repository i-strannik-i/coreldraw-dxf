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
    # A combined Corel curve can contain repeated subpaths even when the object
    # as a whole has no identical copy. Do not count whole-object copies twice.
    whole_copies={(group['layer'],sid) for group in groups for sid in group['copies']}
    paths=defaultdict(list)
    ordinals=defaultdict(int)
    for contour in contours:
        sid=contour['id'];layer=contour['layer']
        index=contour.get('subpath',ordinals[(layer,sid)])
        ordinals[(layer,sid)]+=1
        if (layer,sid) in whole_copies or not contour['segments']:continue
        paths[(layer,contour_key(contour))].append((sid,index,contour))
    for (layer,_),parts in paths.items():
        if len(parts)<2:continue
        parts.sort(key=lambda part:(part[0],part[1]))
        members=[{'id':sid,'subpath':index} for sid,index,_ in parts]
        groups.append(dict(kind='duplicate',mode='subpaths',layer=layer,
            ids=sorted({sid for sid,_,_ in parts}),keep=parts[0][0],
            copies=[sid for sid,_,_ in parts[1:]],members=members,
            point=parts[0][2]['segments'][0][1][0],approx=False))
    return groups


def duplicate_removals(groups):
    """Whole object IDs and per-object subpath indexes to quarantine."""
    whole=set();partial=defaultdict(set)
    for group in groups:
        if group.get('mode')=='subpaths':
            for member in group['members'][1:]:
                partial[member['id']].add(member['subpath'])
        else:whole.update(group['copies'])
    return whole,dict(partial)
