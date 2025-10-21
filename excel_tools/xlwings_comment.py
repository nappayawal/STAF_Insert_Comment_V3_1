"""
xlwings/COM helpers to insert legacy Notes (comments) safely without deleting shapes.

IMPORTANT:
- Requires Excel installed (Windows)
- Uses legacy Note (Range.AddComment) to keep formatting simple and robust
- Saves to *_with_Note.xlsm by default (macro-enabled, FileFormat=52)
"""
from pathlib import Path
import os
import xlwings as xw

def insert_comment_at_address(
    in_path: str,
    sheet_name: str,
    cell: str,
    note_text: str,
    out_path: str | None = None,
    make_visible: bool = False,
    autosize: bool = True,
    width: float | None = None,
    height: float | None = None,
) -> dict:
    """
    Insert or update a legacy Note (old 'Comment') at a specific cell.
    - Avoids duplicates (skips if same content already exists)
    - Preserves shapes/graphics by letting Excel handle the write via COM
    - Returns a summary incl. shapes count before/after

    NOTE: If out_path is None, will save next to input as *_with_Note.xlsm
    """
    in_path = str(Path(in_path).expanduser())
    if out_path is None:
        p = Path(in_path)
        out_path = str(p.with_name(p.stem + "_with_Note" + p.suffix))

    app = None
    wb = None
    created = False
    updated = False
    skipped = False
    shapes_before = None
    shapes_after = None

    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(in_path, update_links=False, read_only=False)
        ws = wb.sheets[sheet_name]

        ws_api = ws.api
        rng_api = ws.range(cell).api

        # Count shapes before change
        shapes_before = int(ws_api.Shapes.Count)

        # Legacy comment handling
        existing = rng_api.Comment  # None or a Comment COM object
        if existing is not None:
            existing_text = existing.Text()
            if (existing_text or "").strip() == (note_text or "").strip():
                skipped = True  # identical -> don't duplicate
            else:
                rng_api.ClearComments()
                #rng_api.AddComment(note_text)
                rng_api.AddComment(note_text.replace("\\n", "\n"))

                updated = True
        else:
            #rng_api.AddComment(note_text)
            rng_api.AddComment(note_text.replace("\\n", "\n"))

            created = True

        # Format size/visibility
        if rng_api.Comment is not None:
            comment = rng_api.Comment
            comment.Visible = bool(make_visible)
            shp = comment.Shape
            if autosize:
                shp.Width = 200
                shp.Height = 100
                shp.TextFrame.AutoSize = True
            else:
                if width is not None:
                    shp.Width = float(width)
                if height is not None:
                    shp.Height = float(height)

        shapes_after = int(ws_api.Shapes.Count)

        # Save with correct format (52 = xlsm)
        if os.path.abspath(out_path) != os.path.abspath(in_path):
            wb.api.SaveAs(out_path, FileFormat=52)
        else:
            wb.save()

        return {
            "in_path": in_path,
            "out_path": out_path,
            "sheet": sheet_name,
            "cell": cell,
            "created": created,
            "updated": updated,
            "skipped": skipped,
            "shapes_before": shapes_before,
            "shapes_after": shapes_after,
            "shapes_intact": (shapes_before == shapes_after),
        }
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if app is not None:
            try:
                app.kill()
            except Exception:
                pass


def insert_comments_batch(
    in_path: str,
    sheet_name: str,
    placements: list[tuple[str, str]],   # [(cell_addr, note_text), ...]
    out_path: str | None = None,
    make_visible: bool = False,
    autosize: bool = True,
) -> dict:
    """
    Batch insert/update legacy Notes (old comments) in a single Excel session.
    - placements: list of (cell_address, note_text)
    - Saves as *_with_Note.xlsm by default (FileFormat=52)
    - Returns a summary (created/updated/skipped counts + shapes check)
    """
    in_path = str(Path(in_path).expanduser())
    if out_path is None:
        p = Path(in_path)
        out_path = str(p.with_name(p.stem + "_with_Note" + p.suffix))

    app = wb = None
    created = updated = skipped = 0
    shapes_before = shapes_after = None

    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(in_path, update_links=False, read_only=False)
        ws = wb.sheets[sheet_name]
        ws_api = ws.api

        shapes_before = int(ws_api.Shapes.Count)

        for cell, note_text in placements:
            rng_api = ws.range(cell).api
            existing = rng_api.Comment
            if existing is not None:
                existing_text = existing.Text()
                if (existing_text or "").strip() == (note_text or "").strip():
                    skipped += 1
                    continue
                rng_api.ClearComments()
                #rng_api.AddComment(note_text)
                rng_api.AddComment(note_text.replace("\\n", "\n"))

                updated += 1
            else:
                #rng_api.AddComment(note_text)
                rng_api.AddComment(note_text.replace("\\n", "\n"))

                created += 1

            # sizing/visibility per note
            if rng_api.Comment is not None:
                cmt = rng_api.Comment
                cmt.Visible = bool(make_visible)
                shp = cmt.Shape
                if autosize:
                    shp.Width = 200
                    shp.Height = 100
                    shp.TextFrame.AutoSize = True

        shapes_after = int(ws_api.Shapes.Count)

        # Save with correct format (52 = .xlsm)
        if os.path.abspath(out_path) != os.path.abspath(in_path):
            wb.api.SaveAs(out_path, FileFormat=52)
        else:
            wb.save()

        return {
            "in_path": in_path,
            "out_path": out_path,
            "sheet": sheet_name,
            "placements": len(placements),
            "created": created,
            "updated": updated,
            "skipped": skipped,
            "shapes_before": shapes_before,
            "shapes_after": shapes_after,
            "shapes_intact": (shapes_before == shapes_after),
        }
    finally:
        if wb is not None:
            try: wb.close()
            except Exception: pass
        if app is not None:
            try: app.kill()
            except Exception: pass
