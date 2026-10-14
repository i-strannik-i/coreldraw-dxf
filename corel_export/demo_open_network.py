"""Verify the user's tail-first, no-repeat traversal on the crossing demo."""
from pathlib import Path
import json
import math
from demo_network_route import example, closed_network, decode_subpath


def calculate():
    plan=closed_network(example(),.1)
    edges=plan['edges']
    # This example has exactly one permissible gap. Include it once.
    gaps=[e for e in edges if e['source'] is None]
    assert len(gaps)==1 and gaps[0]['weight']<=.1
    adjacency={}
    for i,e in enumerate(edges):
        adjacency.setdefault(e['u'],[]).append((i,e['v']))
        adjacency.setdefault(e['v'],[]).append((i,e['u']))
    odd=[v for v,a in adjacency.items() if len(a)%2]
    assert len(odd)==2
    start=next(v for v in odd if len(adjacency[v])==1)
    used=set();stack=[(start,None)];walk=[]
    while stack:
        u,arrived=stack[-1]
        while adjacency[u] and adjacency[u][-1][0] in used:adjacency[u].pop()
        if adjacency[u]:
            i,v=adjacency[u].pop();used.add(i);stack.append((v,(i,u,v)))
        else:
            stack.pop()
            if arrived is not None:walk.append(arrived)
    walk.reverse()
    assert len(walk)==len(edges)==len({i for i,u,v in walk})
    segments=[]
    for i,u,v in walk:
        e=edges[i];kind,p=e['segment']
        segments.append((kind,list(p) if e['u']==u else list(reversed(p))))
    assert all(math.dist(a[1][-1],b[1][0])<1e-6 for a,b in zip(segments,segments[1:]))
    assert math.dist(segments[0][1][0],segments[-1][1][-1])>1
    report=dict(source_mm=plan['source_length'],length_mm=sum(e['weight'] for e in edges),
                repeated_edges=0,segments=len(segments),connector_mm=gaps[0]['weight'],
                start=segments[0][1][0],end=segments[-1][1][-1])
    return segments,report


def main():
    import win32com.client
    segments,report=calculate()
    app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
    doc=app.CreateDocument();doc.Unit=3;doc.ActivePage.SetSize(210,210)
    layer=doc.ActivePage.CreateLayer('Один открытый вектор — без повторов')
    curve=app.CreateCurve(doc)
    def shift(p):return p[0]+105,p[1]+115
    path=curve.CreateSubPath(*shift(segments[0][1][0]))
    for kind,p in segments:
        if kind=='L':path.AppendLineSegment(*shift(p[-1]))
        else:path.AppendCurveSegment2(*shift(p[-1]),*shift(p[1]),*shift(p[2]))
    shape=layer.CreateCurve(curve);shape.Fill.ApplyNoFill()
    shape.Outline.Width=.35;shape.Outline.Color.RGBAssign(30,130,75)
    assert shape.Curve.SubPaths.Count==1
    path=shape.Curve.SubPaths.Item(1)
    assert not path.Closed
    actual=decode_subpath(path.GetCurveInfo(),1,path.Segments.Count,False)
    assert len(actual)==len(segments)
    for (kind,p),(akind,q) in zip(segments,actual):
        assert kind==akind
        indices=range(4) if kind=='B' else (0,3)
        assert all(math.dist(shift(p[i]),q[i])<.00015 for i in indices)
    info=doc.ActivePage.CreateLayer('INFO');info.Printable=False
    def text(x,y,s,size=12):info.CreateArtisticText(x,y,s,Font='Arial',Size=size)
    text(15,192,'ОДИН ОТКРЫТЫЙ ВЕКТОР — БЕЗ ПОВТОРОВ',15)
    for p,s,color in ((segments[0][1][0],'Начало',(30,150,60)),(segments[-1][1][-1],'Конец',(220,70,30))):
        x,y=shift(p);mark=info.CreateEllipse2(x,y,1.3,1.3)
        mark.Fill.UniformColor.RGBAssign(*color)
        text(x+3,y,s,10)
    text(15,46,f'Длина: {report["length_mm"]:.2f} мм. Повторных участков: 0.')
    text(15,35,'Соединение: 0,06 мм при допуске 0,1 мм.')
    text(15,24,'Начало на конце хвостика, конец — в месте его примыкания.',10)
    info.Editable=False;doc.ClearSelection()
    doc.ActiveWindow.ActiveView.SetViewArea(0,0,210,210)
    folder=Path(__file__).resolve().parent/'build/network-route-demo'
    output=folder/'Один открытый вектор — без повторов.cdr'
    doc.SaveAs(str(output),None)
    report['corel_verified']=True
    (folder/'open-result.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))
    print(output)


if __name__=='__main__':main()
