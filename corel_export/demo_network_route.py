"""Visible crossing-network examples at two tolerances, on a new document."""
from pathlib import Path
import collections
import json
import math
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from network_route import closed_network,value
from corel_geometry import decode_subpath


def example():
    radius=48
    def arc(a,b):
        k=4/3*math.tan((b-a)/4)
        p=(radius*math.cos(a),radius*math.sin(a));q=(radius*math.cos(b),radius*math.sin(b))
        return ('B',[p,(p[0]-k*radius*math.sin(a),p[1]+k*radius*math.cos(a)),
                     (q[0]+k*radius*math.sin(b),q[1]-k*radius*math.cos(b)),q])
    segments=[arc(i*math.pi/4+(.06/radius if i==0 else 0),(i+1)*math.pi/4) for i in range(8)]
    for cx,cy in ((-19,8),(14,-15)):
        for i in range(4):
            kind,points=arc(i*math.pi/2,(i+1)*math.pi/2)
            segments.append((kind,[(x*.7+cx,y*.7+cy) for x,y in points]))
    p=value(segments[2],.45)
    segments.append(('B',[p,(p[0]+5,p[1]-8),(p[0]-8,p[1]-16),(p[0]-3,p[1]-23)]))
    return [dict(layer='TEST',closed=False,segments=[s]) for s in segments]


def validate(plan):
    visits=collections.Counter(i for r in plan['routes'] for i in r['edge_ids'])
    for i,edge in enumerate(plan['edges']):
        if edge['source'] is not None:assert visits[i]>=1
        elif visits[i]:assert edge['weight']<=plan['tolerance']
    for route in plan['routes']:
        assert math.dist(route['segments'][0][1][0],route['segments'][-1][1][-1])<1e-6
    gaps=[e['weight'] for i,e in enumerate(plan['edges']) if e['source'] is None and visits[i]]
    return dict(tolerance=plan['tolerance'],closed_routes=len(plan['routes']),
                graph_junctions=plan['graph_intersections'],source_mm=plan['source_length'],
                route_mm=sum(r['length'] for r in plan['routes']),
                repeat_mm=sum((visits[i]-1)*e['weight'] for i,e in enumerate(plan['edges']) if e['source'] is not None),
                connectors=len(gaps),max_connector_mm=max(gaps,default=0))


def main():
    import win32com.client
    contours=example();plans=[closed_network(contours,t) for t in (.1,.01)]
    reports=[validate(p) for p in plans]
    assert all(len(p['routes'])==1 for p in plans)
    assert reports[0]['connectors']>reports[1]['connectors'],reports
    folder=Path(__file__).resolve().parent/'build/network-route-demo';folder.mkdir(parents=True,exist_ok=True)
    (folder/'result.json').write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(reports,ensure_ascii=False,indent=2),flush=True)
    app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
    doc=app.CreateDocument();doc.Unit=3;page=doc.ActivePage;page.SetSize(420,297)
    before=page.CreateLayer('ДО — пересекающиеся кривые')
    result_layers=[page.CreateLayer('ПОСЛЕ — допуск '+str(t)) for t in (.1,.01)]
    info=page.CreateLayer('INFO');marks=page.CreateLayer('_DXF_CHECK_DEMO');marks.Printable=False
    def draw(segments,layer,offset,color,closed=False,width=.4):
        def shift(p):return (p[0]+offset[0],p[1]+offset[1])
        curve=app.CreateCurve(doc);path=curve.CreateSubPath(*shift(segments[0][1][0]))
        for kind,p in segments:
            if kind=='L':path.AppendLineSegment(*shift(p[-1]))
            else:path.AppendCurveSegment2(*shift(p[-1]),*shift(p[1]),*shift(p[2]))
        if closed:path.Closed=True
        shape=layer.CreateCurve(curve);shape.Fill.ApplyNoFill();shape.Outline.Width=width;shape.Outline.Color.RGBAssign(*color)
        return shape
    def label(x,y,text,size=12,color=(35,45,55)):
        shape=info.CreateArtisticText(x,y,text,Font='Arial',Size=size);shape.Fill.UniformColor.RGBAssign(*color)
    for c in contours:draw(c['segments'],before,(70,170),(40,95,160))
    for index,(plan,layer,report) in enumerate(zip(plans,result_layers,reports)):
        offset=(210+index*140,170)
        for route in plan['routes']:
            shape=draw(route['segments'],layer,offset,(35,135,85),True)
            path=shape.Curve.SubPaths.Item(1)
            assert shape.Curve.SubPaths.Count==1 and path.Closed
            actual=decode_subpath(path.GetCurveInfo(),1,path.Segments.Count,True)
            # Corel may insert a zero closing segment when the final endpoint is
            # already equal to the start. Reject that before showing a result.
            assert all(any(math.dist(p[0],q)>.00001 for q in p[1:]) for kind,p in actual)
        visits=collections.Counter(i for r in plan['routes'] for i in r['edge_ids'])
        for i,count in visits.items():
            e=plan['edges'][i]
            if e['source'] is not None and count>1:draw([e['segment']],marks,offset,(171,72,150),width=.65)
            if e['source'] is None:
                p=e['segment'][1][0]
                marker=marks.CreateEllipse2(p[0]+offset[0],p[1]+offset[1],1.5,1.5)
                marker.Fill.ApplyNoFill();marker.Outline.Width=.6;marker.Outline.Color.RGBAssign(230,115,25)
        x=150+index*140
        label(x,97,f'Маршрут: {report["route_mm"]:.1f} мм',12)
        label(x,85,f'Повторно: {report["repeat_mm"]:.1f} мм',12)
        label(x,73,f'Новых соединений: {report["connectors"]}',12)
        label(x,61,f'Макс. соединение: {report["max_connector_mm"]:.4f} мм',11)
    label(16,272,'ЗАМКНУТЫЙ МАРШРУТ ЧЕРЕЗ ПЕРЕСЕЧЕНИЯ',19)
    label(18,245,f'ДО: {len(contours)} открытых кривых',14,(40,95,160))
    label(155,245,'ПОСЛЕ: допуск 0,1 мм',14,(35,135,85))
    label(295,245,'ПОСЛЕ: допуск 0,01 мм',14,(35,135,85))
    label(16,97,f'Исходные линии: {reports[0]["source_mm"]:.1f} мм',12)
    label(16,85,f'Узлов ветвления: {reports[0]["graph_junctions"]}',12)
    label(16,73,'Разрыв у правого края: 0,06 мм',11)
    label(16,43,'Зелёный — маршрут. Фиолетовый — повторный проход. Оранжевая метка — короткое новое соединение.',11)
    label(16,31,'Все исходные участки пройдены. Оба результата — по одному замкнутому контуру, без холостых переходов.',11)
    label(16,19,'Меньший допуск запрещает перемычку 0,06 мм: фреза возвращается по существующей линии.',11)
    info.Editable=False;marks.Editable=False;doc.ClearSelection()
    doc.ActiveWindow.ActiveView.SetViewArea(0,0,420,297)
    output=folder/'Пересечения и допуск — до и после.cdr';doc.SaveAs(str(output),None)
    print(str(output),flush=True)


if __name__=='__main__':main()
