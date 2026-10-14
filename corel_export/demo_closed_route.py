"""Create a visible before/after example without modifying the user's drawing."""
from pathlib import Path
import json
import math
import sys
import win32com.client
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'engine'))
from closed_route import shortest_closed_route,oriented_segments
from corel_geometry import decode_subpath
from vector_audit import audit,flatten


def make_piece(index):
    span=math.radians(46)
    def sample(t):
        angle=math.radians(index*60+7)+span*t
        radius=60+8*math.sin(4*math.pi*t)+5*math.sin(2*math.pi*t)
        dr=32*math.pi*math.cos(4*math.pi*t)+10*math.pi*math.cos(2*math.pi*t)
        co,si=math.cos(angle),math.sin(angle)
        return (radius*co,radius*si),(dr*co-radius*si*span,dr*si+radius*co*span)
    segments=[]
    for number in range(8):
        a,da=sample(number/8);b,db=sample((number+1)/8)
        c1=tuple(a[i]+da[i]/24 for i in range(2))
        c2=tuple(b[i]-db[i]/24 for i in range(2))
        segments.append(('B',[a,c1,c2,b]))
    return dict(closed=False,segments=segments)


pieces=[make_piece(i) for i in (0,3,1,5,2,4)]
for i in (1,3,4):pieces[i]['segments']=oriented_segments(pieces[i],True)
plan=shortest_closed_route(pieces)
checked=dict(id=1,layer='OUT',closed=True,segments=plan['segments'])
diagnostic=audit([checked],tolerance=.01)
assert diagnostic['complete'] and not diagnostic['issues'],diagnostic
length=lambda segments:sum(sum(math.dist(a,b) for a,b in zip(points,points[1:]))
    for kind,control in segments for points in [flatten(control,.0001) if kind=='B' else [control[0],control[-1]]])
original_length=sum(length(c['segments']) for c in pieces)
report=dict(open_before=6,closed_after=1,open_after=0,original_beziers=48,
            added_lines=len(plan['bridges']),source_length_mm=original_length,
            arbitrary_added_mm=plan['baseline_added_length'],optimal_added_mm=plan['added_length'],
            total_closed_length_mm=original_length+plan['added_length'],order=plan['order'],
            exact_for_fixed_original_curves=True,intersections=0)
folder=Path(__file__).resolve().parent/'build/closed-route-demo'
folder.mkdir(parents=True,exist_ok=True)
(folder/'result.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
doc=app.CreateDocument();doc.Unit=3;page=doc.ActivePage;page.SetSize(420,297)
before=page.CreateLayer('ДО — 6 открытых кривых')
after=page.CreateLayer('ПОСЛЕ — 1 замкнутый маршрут')
info=page.CreateLayer('INFO')
highlights=page.CreateLayer('_DXF_CHECK_DEMO');highlights.Printable=False

def shifted(point,offset):return (point[0]+offset[0],point[1]+offset[1])
def draw(segments,layer,offset,closed,color):
    curve=app.CreateCurve(doc)
    path=curve.CreateSubPath(*shifted(segments[0][1][0],offset))
    # Corel itself adds the final straight closing segment.
    body=segments[:-1] if closed else segments
    for kind,points in body:
        end=shifted(points[-1],offset)
        if kind=='L':path.AppendLineSegment(*end)
        else:
            c1=shifted(points[1],offset);c2=shifted(points[2],offset)
            path.AppendCurveSegment2(*end,*c1,*c2)
    if closed:path.Closed=True
    shape=layer.CreateCurve(curve);shape.Fill.ApplyNoFill()
    shape.Outline.Width=.45;shape.Outline.Color.RGBAssign(*color)
    return shape

def label(x,y,text,size=12,color=(35,45,55)):
    shape=info.CreateArtisticText(x,y,text,Font='Arial',Size=size)
    shape.Fill.UniformColor.RGBAssign(*color)
    return shape

for number,contour in enumerate(pieces,1):
    draw(contour['segments'],before,(108,160),False,(36,97,169))
    for point in (contour['segments'][0][1][0],contour['segments'][-1][1][-1]):
        x,y=shifted(point,(108,160))
        mark=highlights.CreateEllipse2(x,y,1,1)
        mark.Fill.ApplyNoFill();mark.Outline.Color.RGBAssign(224,104,35);mark.Outline.Width=.35
    point=contour['segments'][0][1][0]
    label(108+point[0]*1.12,160+point[1]*1.12,str(number),10)

doc.BeginCommandGroup('Замкнуть шесть кривых оптимальным маршрутом')
result=draw(plan['segments'],after,(312,160),True,(30,135,82))
doc.EndCommandGroup()
assert result.Curve.SubPaths.Count==1 and result.Curve.SubPaths.Item(1).Closed
path=result.Curve.SubPaths.Item(1)
actual=decode_subpath(path.GetCurveInfo(),1,path.Segments.Count,True)
assert len(actual)==54,len(actual)
expected=[(kind,[shifted(p,(312,160)) for p in points]) for kind,points in plan['segments']]
def same_cycle(actual,expected):
    # Corel stores coordinates at 0.0001 mm and can put the closing segment first.
    for shift in range(len(expected)):
        rotated=expected[shift:]+expected[:shift]
        if all(k==ek and all(math.dist(p,q)<.00015 for p,q in zip(points,ep))
               for (k,points),(ek,ep) in zip(actual,rotated)):return True
    return False
assert same_cycle(actual,expected)
doc.Undo();assert after.Shapes.Count==0
doc.Redo();assert after.Shapes.Count==1
for a,b in plan['bridges']:
    mark=highlights.CreateLineSegment(*shifted(a,(312,160)),*shifted(b,(312,160)))
    mark.Outline.Width=.85;mark.Outline.Color.RGBAssign(224,104,35)

label(20,274,'РЕАЛЬНОЕ ЗАМЫКАНИЕ С ОПТИМИЗАЦИЕЙ МАРШРУТА',19)
label(26,244,'ДО: 6 открытых кривых',16,(36,97,169))
label(235,244,'ПОСЛЕ: 1 замкнутый контур',16,(30,135,82))
label(23,67,'48 исходных криволинейных участков сохранены.',12)
label(23,55,'Оранжевым — 6 добавленных соединений. По ним фреза режет.',12)
label(23,39,f'Соединения в исходном порядке: {plan["baseline_added_length"]:.1f} мм. Оптимальные: {plan["added_length"]:.1f} мм.',13)
label(23,26,f'Полный замкнутый маршрут: {original_length+plan["added_length"]:.1f} мм. Открытых концов: 0. Пересечений: 0.',13)
label(23,13,'Точный минимум для этих 6 кривых: выбран порядок и направление, каждый исходный участок пройден один раз.',10)
info.Editable=False;highlights.Editable=False
doc.ClearSelection()
doc.ActiveWindow.ActiveView.SetViewArea(0,0,420,297)
output=folder/'Замыкание — до и после.cdr'
doc.SaveAs(str(output),None)
print(json.dumps(dict(report,file=str(output)),ensure_ascii=False,indent=2),flush=True)
