"""Read back the actual saved Corel routes and compare every control point."""
import math
import win32com.client
from demo_network_route import example,closed_network,decode_subpath

app=win32com.client.gencache.EnsureDispatch('CorelDRAW.Application.25')
doc=app.ActiveDocument
assert doc.Name=='Пересечения и допуск — до и после.cdr'
for index,tolerance in enumerate((.1,.01)):
    layer=doc.ActivePage.Layers.Item('ПОСЛЕ — допуск '+str(tolerance))
    assert layer.Shapes.Count==1
    shape=layer.Shapes.Item(1)
    assert shape.Curve.SubPaths.Count==1
    path=shape.Curve.SubPaths.Item(1)
    assert path.Closed
    actual=decode_subpath(path.GetCurveInfo(),1,path.Segments.Count,True)
    expected=closed_network(example(),tolerance)['routes'][0]['segments']
    offset=(210+index*140,170)
    expected=[(kind,[(x+offset[0],y+offset[1]) for x,y in points]) for kind,points in expected]
    assert len(actual)==len(expected),(len(actual),len(expected))
    def equal(a,b):
        if a[0]!=b[0]:return False
        points=zip(a[1],b[1]) if a[0]=='B' else zip((a[1][0],a[1][-1]),(b[1][0],b[1][-1]))
        return all(math.dist(p,q)<.00015 for p,q in points)
    assert any(all(equal(actual[i],expected[(i+shift)%len(expected)]) for i in range(len(actual))) for shift in range(len(expected)))
    print(f'Tolerance {tolerance}: one closed Corel curve, {len(actual)} segments; all control points verified.')
