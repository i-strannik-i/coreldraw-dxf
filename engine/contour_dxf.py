"""One connected DXF polyline per source subpath, with bounded approximation."""
import math


def distance_to_chord(point, start, end):
    dx, dy = end[0]-start[0], end[1]-start[1]
    length2 = dx*dx + dy*dy
    t = max(0., min(1., ((point[0]-start[0])*dx + (point[1]-start[1])*dy)/length2)) if length2 else 0.
    return math.hypot(point[0]-start[0]-t*dx, point[1]-start[1]-t*dy)


def polyline_vertices(segments, closed, tolerance):
    from convert_cdr import circle_fit, split
    if tolerance <= 0 or not math.isfinite(tolerance):
        raise ValueError('Для экспорта связанных контуров допуск аппроксимации должен быть больше нуля.')
    vertices = []

    def emit(points, depth=0):
        start, c1, c2, end = points
        # The Bezier lies inside its control hull; a finite-chord bound also
        # rejects backtracking beyond the endpoints, unlike a line-distance test.
        if max(distance_to_chord(p, start, end) for p in (c1, c2)) <= tolerance:
            vertices.append((*start, 0.))
            return
        arc = circle_fit(points, tolerance)
        if arc:
            center, radius, a, b = arc
            cross = (start[0]-center[0])*(c1[1]-start[1])-(start[1]-center[1])*(c1[0]-start[0])
            sweep = math.radians((b-a) % 360)
            vertices.append((*start, math.tan(sweep/4) * (1 if cross > 0 else -1)))
            return
        if depth >= 24:
            raise ValueError('Не удалось представить кривую с заданным допуском DXF.')
        left, right = split(points)
        emit(left, depth+1)
        emit(right, depth+1)

    previous = None
    for kind, points in segments:
        if previous is not None and math.dist(previous, points[0]) > 1e-7:
            raise ValueError('Обнаружен разрыв внутри исходного контура. Итоговый DXF не сохранён.')
        if kind == 'L':
            if math.dist(points[0], points[-1]) > 1e-10:
                vertices.append((*points[0], 0.))
        else:
            emit(points)
        previous = points[-1]
    if not vertices:
        raise ValueError('Исходный контур имеет нулевую длину.')
    if closed:
        if math.dist(segments[0][1][0], previous) > 1e-7:
            raise ValueError('У замкнутого исходного контура не совпадают концы. Итоговый DXF не сохранён.')
    else:
        vertices.append((*previous, 0.))
    if len(vertices) < 2:
        raise ValueError('Вырожденный исходный контур: недостаточно разных точек.')
    return vertices


def requires_closed(layer):
    name = layer.strip().upper()
    if name.startswith('BACK_'):
        name = name[5:]
    return name in ('CUT_OUT', 'CUT_IN', 'OUT', 'IN', 'P', 'D') or name.startswith('P_') or (name.startswith('D') and name[1:].replace('.', '', 1).isdigit())
