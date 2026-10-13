"""Prototype: tolerance-bounded closed routes through a crossing vector network.

Bezier geometry is split, not flattened in the output. Short optional connectors
are allowed only from source endpoints. Exact odd-node matching is bounded to
18 odd vertices per component. No claim of a global optimum over all possible
ways of connecting initially disconnected components is made.
"""
import heapq
import math
from functools import lru_cache
try:
    from .contour_dxf import distance_to_chord
except ImportError:
    from contour_dxf import distance_to_chord


def lerp(a,b,t):return tuple(x+(y-x)*t for x,y in zip(a,b))
def cross(a,b):return a[0]*b[1]-a[1]*b[0]
def sub(a,b):return (a[0]-b[0],a[1]-b[1])
def split(p,t):
    rows=[list(p)]
    while len(rows[-1])>1:rows.append([lerp(a,b,t) for a,b in zip(rows[-1],rows[-1][1:])])
    return [row[0] for row in rows],[row[-1] for row in reversed(rows)]


def value(segment,t):
    kind,p=segment
    if kind=='L':return lerp(p[0],p[-1],t)
    return split(p,t)[0][-1]


def derivative(segment,t):
    kind,p=segment
    if kind=='L':return sub(p[-1],p[0])
    return tuple(3*((1-t)**2*(p[1][k]-p[0][k])+2*t*(1-t)*(p[2][k]-p[1][k])+t*t*(p[3][k]-p[2][k])) for k in (0,1))


def portion(segment,a,b):
    kind,p=segment
    if kind=='L':
        x,y=value(segment,a),value(segment,b)
        return kind,[x,x,y,y]
    if b<1:p=split(p,b)[0]
    if a>0:p=split(p,a/b)[1]
    return kind,p


def samples(segment,error):
    kind,p=segment
    if kind=='L':return [(0,p[0]),(1,p[-1])]
    def visit(p,a,b,depth):
        if max(distance_to_chord(q,p[0],p[-1]) for q in p[1:3])<=error:
            return [(a,p[0]),(b,p[-1])]
        if depth>=22:raise ValueError('Кривая слишком сложна для тестового расчёта.')
        left,right=split(p,.5);middle=(a+b)/2
        return visit(left,a,middle,depth+1)[:-1]+visit(right,middle,b,depth+1)
    return visit(p,0,1,0)


def refine(first,second,t,u,epsilon):
    for _ in range(25):
        p,q=value(first,t),value(second,u);difference=sub(q,p)
        if math.dist(p,q)<=epsilon:return max(0,min(1,t)),max(0,min(1,u))
        a,b=derivative(first,t),derivative(second,u);den=cross(a,b)
        if abs(den)<1e-14:return None
        t+=cross(difference,b)/den;u+=cross(difference,a)/den
        if not -1e-7<=t<=1+1e-7 or not -1e-7<=u<=1+1e-7:return None
    return None


def build_network(contours,tolerance,progress=None):
    report=progress or (lambda *args:None)
    if not math.isfinite(tolerance) or not .001<=tolerance<=1:
        raise ValueError('Допуск должен быть от 0,001 до 1 мм.')
    if len({c.get('layer','') for c in contours})>1:
        raise ValueError('Каждый слой рассчитывается отдельно.')
    epsilon=min(1e-6,tolerance/1000)
    approximation=min(.001,tolerance/20)
    source=[];tips=[]
    for contour in contours:
        base=len(source);source.extend(contour['segments'])
        if not contour['closed'] and contour['segments']:
            tips.extend([(base,0),(len(source)-1,1)])
    if not source:raise ValueError('Нет участков для расчёта.')
    cuts=[{0.0,1.0} for _ in source]
    polylines=[]
    for index,segment in enumerate(source):
        report("Подготовка участков",index,len(source))
        polylines.append(samples(segment,approximation))
    spans=[]
    for index,points in enumerate(polylines):
        for (t,a),(u,b) in zip(points,points[1:]):
            spans.append((min(a[0],b[0]),max(a[0],b[0]),min(a[1],b[1]),max(a[1],b[1]),index,t,u,a,b))
    spans.sort();active=[]
    for span_index,current in enumerate(spans):
        report("Поиск пересечений",span_index,len(spans))
        active=[old for old in active if old[1]+epsilon>=current[0]]
        for old in active:
            if old[3]+epsilon<current[2] or current[3]+epsilon<old[2]:continue
            i,j=old[4],current[4]
            if i==j and abs(old[6]-current[5])<1e-12:continue
            a,b=old[7:];c,d=current[7:]
            r,s=sub(b,a),sub(d,c);den=cross(r,s)
            if abs(den)<1e-14:
                if i!=j and source[i][0]==source[j][0]=='L' and abs(cross(sub(c,a),r))<epsilon*max(math.hypot(*r),1):
                    rr=sum(v*v for v in r)
                    if rr:
                        ts=[sum(x*y for x,y in zip(sub(p,a),r))/rr for p in (c,d)]
                        if min(1,max(ts))-max(0,min(ts))>epsilon:
                            raise ValueError('Сначала устраните наложения прямых участков.')
                continue
            v=cross(sub(c,a),s)/den;w=cross(sub(c,a),r)/den
            if not -1e-8<=v<=1+1e-8 or not -1e-8<=w<=1+1e-8:continue
            t=old[5]+v*(old[6]-old[5]);u=current[5]+w*(current[6]-current[5])
            if i==j and abs(t-u)<1e-7:continue
            found=refine(source[i],source[j],t,u,epsilon/10)
            if found:cuts[i].add(found[0]);cuts[j].add(found[1])
        active.append(current)
    # Endpoint-to-edge connections: project to sampled pieces, then refine the
    # parameter by bounded minimization on that local interval. Confirm the
    # actual Bezier distance before accepting a connector.
    # Index small sampled spans rather than testing every endpoint against the
    # whole drawing. Exact Bezier refinement still decides each connection.
    extent=max(max(s[1] for s in spans)-min(s[0] for s in spans),
               max(s[3] for s in spans)-min(s[2] for s in spans))
    cell=max(tolerance*4,extent/max(1,math.sqrt(len(spans))),.01)
    buckets={}
    for span in spans:
        for x in range(math.floor((span[0]-tolerance-approximation)/cell),math.floor((span[1]+tolerance+approximation)/cell)+1):
            for y in range(math.floor((span[2]-tolerance-approximation)/cell),math.floor((span[3]+tolerance+approximation)/cell)+1):
                buckets.setdefault((x,y),[]).append(span)
    gaps=[]
    for tip_index,(si,st) in enumerate(tips):
        report("Поиск соединений в допуске",tip_index,len(tips))
        point=value(source[si],st)
        candidates={}
        for span in buckets.get((math.floor(point[0]/cell),math.floor(point[1]/cell)),[]):
            if span[4]!=si:candidates.setdefault(span[4],[]).append(span)
        if epsilon<math.dist(value(source[si],0),value(source[si],1))<=tolerance:
            gaps.append((si,0,si,1))
        for target,local_spans in candidates.items():
            best=None
            for span in local_spans:
                ta,tb,a,b=span[5:]
                if distance_to_chord(point,a,b)>tolerance+approximation:continue
                lo,hi=ta,tb
                for _ in range(35):
                    x=lo+(hi-lo)/3;y=hi-(hi-lo)/3
                    if math.dist(point,value(source[target],x))<math.dist(point,value(source[target],y)):hi=y
                    else:lo=x
                for t in (ta,tb,(lo+hi)/2):
                    distance=math.dist(point,value(source[target],t))
                    if best is None or distance<best[0]:best=(distance,t)
            if best is not None and best[0]<=tolerance:
                cuts[target].add(best[1])
                if best[0]>epsilon:gaps.append((si,st,target,best[1]))
    vertices=[];cells={}
    def vertex(point):
        key=(math.floor(point[0]/epsilon),math.floor(point[1]/epsilon))
        for dx in (-1,0,1):
            for dy in (-1,0,1):
                for index in cells.get((key[0]+dx,key[1]+dy),[]):
                    if math.dist(point,vertices[index])<=epsilon:return index
        index=len(vertices);vertices.append(point);cells.setdefault(key,[]).append(index);return index
    edges=[]
    def add_edge(segment,origin,required):
        kind,p=segment;p=list(p);u,v=vertex(p[0]),vertex(p[-1])
        p[0],p[-1]=vertices[u],vertices[v]
        if kind=='L':p=[p[0],p[0],p[-1],p[-1]]
        samples_=samples((kind,p),min(.0001,tolerance/100))
        weight=sum(math.dist(a,b) for (_,a),(_,b) in zip(samples_,samples_[1:]))
        if weight<=epsilon:return
        edges.append(dict(u=u,v=v,segment=(kind,p),weight=weight,source=origin,required=required))
    for index,segment in enumerate(source):
        report("Разделение кривых на пересечениях",index,len(source))
        parameters=sorted(cuts[index]);unique=[parameters[0]]
        for t in parameters[1:]:
            if t-unique[-1]>1e-9:unique.append(t)
        for a,b in zip(unique,unique[1:]):add_edge(portion(segment,a,b),(index,a,b),True)
    gap_pairs=set()
    for si,st,ti,tt in gaps:
        a,b=value(source[si],st),value(source[ti],tt)
        u,v=vertex(a),vertex(b);pair=tuple(sorted((u,v)))
        # Recheck after snapping to graph coordinates; no connector may exceed
        # the user's actual tolerance.
        if u==v or pair in gap_pairs or math.dist(vertices[u],vertices[v])>tolerance:continue
        gap_pairs.add(pair);add_edge(('L',[a,a,b,b]),None,False)
    return dict(edges=edges,vertices=vertices,tolerance=tolerance)


def closed_network(contours,tolerance):
    graph=build_network(contours,tolerance)
    edges,vertices=graph['edges'],graph['vertices']
    epsilon=min(1e-6,tolerance/1000)
    parent=list(range(len(vertices)))
    def root(u):
        while parent[u]!=u:parent[u]=parent[parent[u]];u=parent[u]
        return u
    def unite(u,v):parent[root(u)]=root(v)
    for edge in edges:
        if edge['required']:unite(edge['u'],edge['v'])
    for edge in sorted((e for e in edges if not e['required']),key=lambda e:e['weight']):
        if root(edge['u'])!=root(edge['v']):edge['required']=True;unite(edge['u'],edge['v'])
    groups={}
    for i,e in enumerate(edges):groups.setdefault(root(e['u']),[]).append(i)
    routes=[]
    for group in groups.values():
        adjacent={};degree={};required=[]
        for index in group:
            e=edges[index];u,v=e['u'],e['v']
            adjacent.setdefault(u,[]).append((v,index));adjacent.setdefault(v,[]).append((u,index))
            if e['required']:
                required.append(index);degree[u]=degree.get(u,0)+1;degree[v]=degree.get(v,0)+1
        if not required:continue
        odd=sorted(v for v,d in degree.items() if d%2)
        if len(odd)>18:raise ValueError('В тестовой сети больше 18 нечётных узлов. Выделите меньшую область.')
        distances={};paths={}
        for start in odd:
            distance={start:0};previous={};heap=[(0,start)]
            while heap:
                cost,u=heapq.heappop(heap)
                if cost!=distance[u]:continue
                for v,index in adjacent[u]:
                    new=cost+edges[index]['weight']
                    if new<distance.get(v,math.inf):
                        distance[v]=new;previous[v]=(u,index);heapq.heappush(heap,(new,v))
            for end in odd:
                if start>=end:continue
                sequence=[];u=end
                while u!=start:u,index=previous[u];sequence.append(index)
                distances[start,end]=distance[end];paths[start,end]=sequence
        @lru_cache(None)
        def matching(mask):
            if not mask:return 0,()
            i=(mask&-mask).bit_length()-1;remaining=mask^(1<<i);best=(math.inf,())
            for j in range(i+1,len(odd)):
                if not remaining&(1<<j):continue
                cost,pairs=matching(remaining^(1<<j));pair=(odd[i],odd[j])
                candidate=(cost+distances[pair],(pair,)+pairs)
                if candidate[0]<best[0]:best=candidate
            return best
        extra,pairs=matching((1<<len(odd))-1)
        occurrences=list(required)
        for pair in pairs:occurrences.extend(paths[pair])
        adjacency={}
        for occurrence,index in enumerate(occurrences):
            e=edges[index]
            adjacency.setdefault(e['u'],[]).append((occurrence,e['v']))
            adjacency.setdefault(e['v'],[]).append((occurrence,e['u']))
        used=set();stack=[(edges[required[0]]['u'],None)];walk=[]
        while stack:
            u,arrived=stack[-1]
            while adjacency[u] and adjacency[u][-1][0] in used:adjacency[u].pop()
            if adjacency[u]:
                occurrence,v=adjacency[u].pop();used.add(occurrence);stack.append((v,(occurrence,u,v)))
            else:
                stack.pop()
                if arrived is not None:walk.append(arrived)
        walk.reverse();segments=[];edge_ids=[]
        for occurrence,u,v in walk:
            index=occurrences[occurrence];edge_ids.append(index);e=edges[index];kind,p=e['segment']
            segments.append((kind,list(p) if e['u']==u else list(reversed(p))))
        assert len(walk)==len(occurrences)
        assert all(math.dist(a[1][-1],b[1][0])<=epsilon for a,b in zip(segments,segments[1:]+segments[:1]))
        routes.append(dict(segments=segments,edge_ids=edge_ids,odd_vertices=len(odd),
                           length=sum(edges[i]['weight'] for i in occurrences),extra_length=extra,
                           required_length=sum(edges[i]['weight'] for i in required)))
    return dict(routes=routes,edges=edges,vertices=vertices,tolerance=tolerance,
                source_length=sum(e['weight'] for e in edges if e['source'] is not None),
                graph_intersections=sum(1 for v in range(len(vertices)) if sum(e['u']==v for e in edges if e['source'] is not None)+sum(e['v']==v for e in edges if e['source'] is not None)>=3))
