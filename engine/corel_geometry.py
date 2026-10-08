"""Decode Corel's bulk CurveElement array without per-node COM round trips."""
import math


def decode_subpath(records, scale, expected_count, closed):
    if not math.isfinite(scale) or scale <= 0:
        raise ValueError('Invalid Corel coordinate scale')
    segments, controls = [], []
    first = current = None
    for record in records:
        point = (record.PositionX * scale, record.PositionY * scale)
        if not all(math.isfinite(value) for value in point):
            raise ValueError('Non-finite Corel curve coordinates')
        kind = record.ElementType
        if kind == 0:
            if first is not None:
                raise ValueError('Multiple starts in a single Corel subpath')
            first = current = point
        elif current is None:
            raise ValueError('Corel subpath has no start')
        elif kind == 3:
            controls.append(point)
            if len(controls) > 2:
                raise ValueError('Too many Corel Bezier controls')
        elif kind == 1:
            if controls:
                raise ValueError('Unexpected controls before Corel line')
            segments.append(('L', [current, current, point, point]))
            current = point
        elif kind == 2:
            if len(controls) != 2:
                raise ValueError('Corel Bezier requires two controls')
            segments.append(('B', [current, *controls, point]))
            controls = []
            current = point
        else:
            raise ValueError('Unknown Corel curve element: ' + str(kind))
    if controls or not segments or len(segments) != expected_count:
        raise ValueError('Incomplete Corel bulk subpath')
    # GetCurveInfo includes the closing segment; never invent or drop one.
    if closed and math.dist(first, current) > 1e-6:
        raise ValueError('Corel bulk subpath is not closed')
    return segments
