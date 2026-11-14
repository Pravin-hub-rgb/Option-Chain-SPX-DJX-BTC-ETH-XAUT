import os


def long_form(path):
    """The Windows long form of a path, or the path itself.

    Mirror of the Delta tool's fix: Workbooks.Open may have been handed the
    8.3 short name, and Excel then reports THAT form as the book's fullname.
    Comparing short-vs-long never matches, so the tool would sit unable to
    write even though the workbook is wide open.
    """
    try:
        import ctypes

        buffer = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.GetLongPathNameW(path, buffer, 1024):
            return buffer.value or path
    except Exception:
        pass
    return path


def find_book_and_app_in_apps(apps, filename, on_error=None):
    target = os.path.normcase(long_form(os.path.abspath(filename)))
    for app in apps:
        try:
            books = list(app.books)
        except Exception as exc:
            if on_error is not None:
                on_error(exc, "enumerating open Excel workbooks")
            continue
        for book in books:
            try:
                if os.path.normcase(long_form(os.path.abspath(book.fullname))) == target:
                    return book, app
            except Exception as exc:
                if on_error is not None:
                    on_error(exc, "reading open workbook path")
                continue
    return None, None


def workbook_name_is_open_in_apps(apps, filename, on_error=None):
    target = os.path.normcase(os.path.abspath(filename))
    for app in apps:
        try:
            books = list(app.books)
        except Exception as exc:
            if on_error is not None:
                on_error(exc, "enumerating workbooks before reopen")
            return True
        for book in books:
            try:
                if os.path.normcase(os.path.abspath(book.fullname)) == target:
                    return True
            except Exception as exc:
                if on_error is not None:
                    on_error(exc, "reading workbook path before reopen")
                return True
    return False


def refresh_delay(configured_refresh, requested_refresh, render_seconds):
    return max(float(configured_refresh), float(requested_refresh), float(render_seconds))


def prefer_expiry_root(rows, expiry, root):
    filtered = [row for row in rows if row.get("expiry") == expiry]
    if any(row.get("root") == root for row in filtered):
        return [row for row in filtered if row.get("root") == root]
    return filtered


def reopen_due(now, last_reopen, interval=20.0):
    return float(now) - float(last_reopen) >= float(interval)


def should_restore_calculation(current_mode, manual_mode):
    return int(current_mode) == int(manual_mode)


def find_book_by_name_in_apps(apps, filename, on_error=None):
    """Return (book, app) for any open workbook with the same file name.

    Excel refuses to open two workbooks with the same name, whatever folder
    they are in. When the customer already has spx_option_chain.xlsx open from
    another directory, the tool cannot open its own copy and has to use the
    one that is already open instead.
    """
    target = os.path.basename(os.path.normcase(os.path.abspath(filename))).lower()
    for app in apps:
        try:
            books = list(app.books)
        except Exception as exc:
            if on_error is not None:
                on_error(exc, "enumerating open Excel workbooks")
            # Cannot inspect this app, so we cannot rule out a duplicate.
            # Say "maybe open" and let the caller leave it alone.
            return None, None
        for book in books:
            try:
                name = os.path.basename(
                    os.path.normcase(os.path.abspath(book.fullname))
                ).lower()
            except Exception as exc:
                if on_error is not None:
                    on_error(exc, "reading open workbook path")
                return None, None
            if name == target:
                return book, app
    return None, None


def workbook_name_open_anywhere_in_apps(apps, filename, on_error=None):
    """True if a workbook with this file name is open in any Excel app.

    Used before asking Excel to open a copy of its own: if a workbook with the
    same name is already open, Excel puts up a modal "cannot open two
    workbooks with the same name" dialog, and the tool used to trigger that
    dialog over and over.

    If Excel cannot be inspected the answer is True, so the tool stays quiet
    rather than risking a duplicate.
    """
    book, _app = find_book_by_name_in_apps(apps, filename, on_error)
    if book is not None:
        return True
    if not apps:
        return False
    # A None result with a non-empty app list means we could not read it.
    return not all(_readable(app) for app in apps)


def _readable(app):
    try:
        list(app.books)
    except Exception:
        return False
    return True
