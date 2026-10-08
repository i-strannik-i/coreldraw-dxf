"""Merge original cubic spans only after a continuous error-bound check."""
import math
import numpy as np


def split_at(p, t):
    a = np.asarray(p, dtype=float)
    levels = [a]
    for _ in range(3):
        a = a[:-1] * (1-t) + a[1:] * t
        levels.append(a)
    return np.array([a[0] for a in levels]), np.array([a[-1] for a in levels[::-1]])


def restrict(p, start, end):
    left = split_at(p, end)[0]
    return split_at(left, start/end)[1] if start else left


def bounded_difference(a, b, tolerance, depth=0):
    delta = np.asarray(a)-np.asarray(b)
    if np.max(np.linalg.norm(delta, axis=1)) <= tolerance:
        return True
    if max(np.linalg.norm(delta[0]), np.linalg.norm(delta[-1])) > tolerance or depth >= 10:
        return False
    left, right = split_at(delta, .5)
    zero = np.zeros((4, 2))
    return bounded_difference(left, zero, tolerance, depth+1) and bounded_difference(right, zero, tolerance, depth+1)


def tangent(p, end=False):
    p = np.asarray(p)
    origin = p[-1] if end else p[0]
    for q in (p[-2::-1] if end else p[1:]):
        d = origin-q if end else q-origin
        length = np.linalg.norm(d)
        if length > 1e-10:
            return d/length
    return np.zeros(2)


def fit(spans, tolerance):
    lengths = [sum(math.dist(a,b) for a,b in zip(p,p[1:])) for p in spans]
    if min(lengths) <= 1e-10:
        return None
    knots = np.r_[0, np.cumsum(lengths)] / sum(lengths)
    start, end = np.asarray(spans[0][0]), np.asarray(spans[-1][-1])
    if np.linalg.norm(end-start) < 1e-9:
        return None
    first, last = tangent(spans[0]), tangent(spans[-1], True)
    matrix, values = [], []
    for i, span in enumerate(spans):
        for u in (0, .25, .5, .75, 1):
            t = knots[i] + u*(knots[i+1]-knots[i])
            weights = np.array([(1-t)**3, 3*t*(1-t)**2, 3*t*t*(1-t), t**3])
            point = split_at(span, u)[0][-1]
            base = (weights[0]+weights[1])*start + (weights[2]+weights[3])*end
            matrix.extend(np.column_stack((weights[1]*first, -weights[2]*last)))
            values.extend(point-base)
    handles = np.linalg.lstsq(matrix, values, rcond=None)[0]
    if not np.all(np.isfinite(handles)) or min(handles) < 0:
        return None
    candidate = np.array([start, start+handles[0]*first, end-handles[1]*last, end])
    if all(bounded_difference(span, restrict(candidate, knots[i], knots[i+1]), tolerance)
           for i, span in enumerate(spans)):
        return candidate.tolist()
    return None


def optimize(segments, tolerance, progress=None):
    """Input/output (kind, cubic controls); original lines stay exact lines."""
    if tolerance <= 0:
        return segments
    result, i = [], 0
    while i < len(segments):
        if progress:
            progress(i, len(segments))
        kind, points = segments[i]
        stop = i+1
        while stop < min(len(segments), i+64) and segments[stop][0] == kind:
            prev, following = segments[stop-1][1], segments[stop][1]
            if math.dist(prev[-1], following[0]) > 1e-8:
                break
            if np.dot(tangent(prev, True), tangent(following)) < math.cos(math.radians(30)):
                break
            if kind == 'L':
                # Only collinear, forward lines are merged; no line becomes a curve.
                direction = np.asarray(following[-1])-points[0]
                offset = np.asarray(prev[-1])-points[0]
                if abs(direction[0]*offset[1]-direction[1]*offset[0]) > 1e-10:
                    break
            stop += 1
        if kind == 'L':
            end = segments[stop-1][1][-1]
            result.append(('L', [points[0], points[0], end, end]))
            i = stop
            continue
        best, count = points, 1
        size = 2
        while size <= stop-i:
            candidate = fit([p for _,p in segments[i:i+size]], tolerance)
            if candidate is None:
                break
            best, count = candidate, size
            size *= 2
        # Search the remaining interval, always certify against ORIGINAL segments.
        low, high = count+1, min(size-1, stop-i)
        while low <= high:
            mid = (low+high)//2
            candidate = fit([p for _,p in segments[i:i+mid]], tolerance)
            if candidate is None:
                high = mid-1
            else:
                best, count, low = candidate, mid, mid+1
        result.append(('B', best))
        i += count
    return result
