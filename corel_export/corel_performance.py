"""Restore Corel's redraw mode on success, cancellation, and COM errors."""
from contextlib import contextmanager


@contextmanager
def paused_redraw(app):
    previous = app.Optimization
    app.Optimization = True
    try:
        yield
    finally:
        app.Optimization = previous
        if not previous:
            app.Refresh()
