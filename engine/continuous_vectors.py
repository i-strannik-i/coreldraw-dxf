"""Edge-disjoint trails of a vector network; never repeat a cutting edge.

For the selected network the trail count is minimal: max(1, odd/2) per
component. Optional tolerance connectors are chosen conservatively, by local
reduction in trail count, then length; that selection is not a global optimum.
"""
import math
try:
    from .network_route import build_network
except ImportError:
    from network_route import build_network


def continuous_vectors(contours,tolerance,coordinate_scale=None):
    seen=set()
    for contour in contours:
        for kind,p in contour['segments']:
            points=p if kind=='B' else (p[0],p[-1])
            key=tuple(tuple(round(x,9) for x in point) for point in points)
            key=(kind,min(key,tuple(reversed(key))))
            if key in seen:raise ValueError('Есть совпадающие участки. Сначала выполните проверку дублей и наложений.')
            seen.add(key)
    graph=build_network(contours,tolerance)
    edges,vertices=graph['edges'],graph['vertices']
    degree=[0]*len(vertices);parent=list(range(len(vertices)))
    def root(u):
        while parent[u]!=u:parent[u]=parent[parent[u]];u=parent[u]
        return u
    selected=[]
    for i,e in enumerate(edges):
        if e['source'] is not None:
            selected.append(i);degree[e['u']]+=1;degree[e['v']]+=1
            parent[root(e['u'])]=root(e['v'])
    odd={}
    for v,d in enumerate(degree):odd[root(v)]=odd.get(root(v),0)+(d%2)
    remaining=[]
    for i,e in enumerate(edges):
        if e['source'] is not None:continue
        if coordinate_scale is not None:
            points=[tuple(round(v/coordinate_scale,4)*coordinate_scale for v in p) for p in (e['segment'][1][0],e['segment'][1][-1])]
            e['weight']=math.dist(*points)
            if e['weight']>tolerance:continue
        remaining.append(i)
    remaining.sort(key=lambda i:edges[i]['weight'])
    def benefit(i):
        e=edges[i];u,v=e['u'],e['v'];a,b=root(u),root(v)
        delta=2-2*(degree[u]%2+degree[v]%2)
        if a==b:return (max(1,odd[a]//2)-max(1,(odd[a]+delta)//2),-delta)
        return (max(1,odd[a]//2)+max(1,odd[b]//2)-max(1,(odd[a]+odd[b]+delta)//2),-delta)
    while remaining:
        best=max(remaining,key=lambda i:(benefit(i),-edges[i]['weight'],-i))
        if benefit(best)<=(0,0):break
        remaining.remove(best);selected.append(best)
        e=edges[best];u,v=e['u'],e['v'];a,b=root(u),root(v)
        count=odd[a]+(odd[b] if a!=b else 0)+2-2*(degree[u]%2+degree[v]%2)
        parent[a]=b;odd[b]=count;degree[u]+=1;degree[v]+=1
    groups={}
    for i in selected:groups.setdefault(root(edges[i]['u']),[]).append(i)
    routes=[];minimum=0
    for group in groups.values():
        adjacency={}
        def add(i,u,v):
            adjacency.setdefault(u,[]).append((i,v));adjacency.setdefault(v,[]).append((i,u))
        for i in group:add(i,edges[i]['u'],edges[i]['v'])
        odds=[v for v in adjacency if len(adjacency[v])%2]
        minimum+=max(1,len(odds)//2)
        # A virtual hub makes an Euler cycle. Removing its virtual edges splits
        # it into the minimum number of open trails, without copying any edge.
        hub=-1
        for n,v in enumerate(odds):add(-n-1,hub,v)
        start=hub if odds else edges[group[0]]['u']
        stack=[(start,None)];walk=[];used=set()
        while stack:
            u,arrival=stack[-1]
            while adjacency[u] and adjacency[u][-1][0] in used:adjacency[u].pop()
            if adjacency[u]:
                i,v=adjacency[u].pop();used.add(i);stack.append((v,(i,u,v)))
            else:
                stack.pop()
                if arrival is not None:walk.append(arrival)
        walk.reverse();pieces=[];piece=[]
        for i,u,v in walk:
            if i<0:
                if piece:pieces.append(piece);piece=[]
            else:piece.append((i,u,v))
        if piece:pieces.append(piece)
        for piece in pieces:
            # Prefer the free end as the start of an open path.
            if degree[piece[-1][2]]==1 and degree[piece[0][1]]!=1:
                piece=[(i,v,u) for i,u,v in reversed(piece)]
            segments=[]
            for i,u,v in piece:
                e=edges[i];kind,p=e['segment']
                segments.append((kind,list(p) if e['u']==u else list(reversed(p))))
            assert all(math.dist(a[1][-1],b[1][0])<1e-5 for a,b in zip(segments,segments[1:]))
            routes.append(dict(closed=piece[0][1]==piece[-1][2],segments=segments,edge_ids=[i for i,u,v in piece]))
    visits=[i for r in routes for i in r['edge_ids']]
    assert len(visits)==len(set(visits))==len(selected) and set(visits)==set(selected)
    assert len(routes)==minimum
    connectors=[edges[i]['weight'] for i in selected if edges[i]['source'] is None]
    graph.update(routes=routes,connectors=len(connectors),max_connector=max(connectors,default=0),
                 minimum_routes=minimum,repeated_edges=0)
    return graph


def continuous_summary(before,after,plan):
    return (f'Векторов: {sum(before)} → {sum(after)}.\n'
            f'Сейчас: замкнутых — {before[0]}, открытых — {before[1]}.\n'
            f'После: замкнутых — {after[0]}, открытых — {after[1]}.\n'
            f'Новых соединений: {plan["connectors"]}; максимальное — {plan["max_connector"]:.4f} мм.\n'
            'Повторных участков: 0. Открытые векторы не замыкаются принудительно.')


def corel_segments(segments,scale):
    """Round once to Corel's four decimal unit coordinates before appending.

    Intersection splits below that resolution can otherwise create new zero
    segments. Only pieces that become a single representable point are omitted.
    """
    result=[]
    for kind,points in segments:
        points=[tuple(round(v/scale,4)*scale for v in p) for p in points]
        if all(p==points[0] for p in points[1:]):continue
        result.append((kind,points))
    if not result:raise ValueError('Вектор меньше точности координат CorelDRAW. Изменения не применены.')
    return result
