# Page Summary - reports for Microsoft Word, Excel and PowerPoint via their object models.
# Every measurement is wrapped so one failure never breaks the rest of the report.
import time

import addonHandler

from .report import Report, counts, plural

addonHandler.initTranslation()

PROG_IDS = {
	"word": "Word.Application",
	"excel": "Excel.Application",
	"powerpoint": "PowerPoint.Application",
}


class OfficeError(Exception):
	pass


def _safe(fn, default=None):
	try:
		return fn()
	except Exception:
		return default


def _int(fn, default=0):
	value = _safe(fn)
	try:
		return int(value)
	except (TypeError, ValueError):
		return default


# Window classes that expose the application's object model (the same route NVDA itself uses).
CHILD_CLASSES = {
	"word": ("_WwG",),
	"excel": ("EXCEL7",),
	"powerpoint": ("paneClassDC", "mdiClass", "screenClass"),
}
OBJID_NATIVEOM = -16
_user32 = None


def _childWindows(parent):
	"""[(hwnd, class name, visible)] for every descendant window of parent."""
	import ctypes
	from ctypes import wintypes
	global _user32
	if _user32 is None:
		_user32 = ctypes.WinDLL("user32")  # private instance: do not touch shared argtypes
	u = _user32
	proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
	u.EnumChildWindows.argtypes = [wintypes.HWND, proc, wintypes.LPARAM]
	u.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
	u.IsWindowVisible.argtypes = [wintypes.HWND]
	found = []

	def callback(hwnd, lparam):
		buf = ctypes.create_unicode_buffer(256)
		u.GetClassNameW(hwnd, buf, 256)
		found.append((hwnd, buf.value, bool(u.IsWindowVisible(hwnd))))
		return True

	cb = proc(callback)
	u.EnumChildWindows(parent, cb, 0)
	return found


def _nativeOM(hwnd):
	"""The Window object of an Office document window, via accessibility (no ROT needed)."""
	import ctypes
	import comtypes.automation
	import comtypes.client
	ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
	ctypes.oledll.oleacc.AccessibleObjectFromWindow(
		hwnd, OBJID_NATIVEOM, ctypes.byref(comtypes.automation.IDispatch._iid_), ctypes.byref(ptr)
	)
	return comtypes.client.dynamic.Dispatch(ptr)


def _rotApp(kind):
	import comtypes.client
	return comtypes.client.GetActiveObject(PROG_IDS[kind], dynamic=True)


def _short(e):
	text = str(e).strip().replace("\n", " ")
	return "%s: %s" % (type(e).__name__, text[:150])


def getApp(kind, hwnd=None, trace=None):
	"""The running Office Application object, or raise OfficeError.

	Tries the document window's object model first (like NVDA does), then the running object table.
	trace, if given, receives a line per attempt for diagnostics.
	"""
	errors = []

	def note(line, error=False):
		if trace is not None:
			trace.append(line)
		if error:
			errors.append(line)

	if hwnd:
		try:
			children = _childWindows(hwnd)
		except Exception as e:
			children = []
			note("Could not list the window's children: " + _short(e), True)
		wanted = CHILD_CLASSES[kind]
		candidates = []
		for visible in (True, False):
			for cls in wanted:
				candidates += [(h, c) for h, c, v in children if c == cls and v == visible]
		if trace is not None:
			note("Candidate windows: " + (", ".join("%s (0x%x)" % (c, h) for h, c in candidates) or "none"))
		for h, cls in candidates:
			try:
				app = _nativeOM(h).Application
				name = app.Name
				note("Connected through the %s window to %s." % (cls, name))
				return app
			except Exception as e:
				note("Window %s: %s" % (cls, _short(e)), True)
	else:
		note("No window handle was available.", True)
	try:
		app = _rotApp(kind)
		note("Connected through the running object table.")
		return app
	except Exception as e:
		note("Running object table: " + _short(e), True)
	reason = errors[-1] if errors else ""
	raise OfficeError(_("Could not connect to the document. %s") % reason)


def _firstDoc(app, attr, protectedAttr):
	"""The active document, or the one in Protected View."""
	doc = _safe(lambda: getattr(app, attr))
	if doc is None:
		doc = _safe(lambda: getattr(app.ActiveProtectedViewWindow, protectedAttr))
	return doc


# ------------------------------------------------------------------ Word

def _headingLevels(items):
	"""GetCrossReferenceItems indents sub-headings; turn indentation into relative levels."""
	indents = [len(s) - len(s.lstrip()) for s in items]
	ranks = {v: i + 1 for i, v in enumerate(sorted(set(indents)))}
	return [(ranks[i], s.strip()) for i, s in zip(indents, items)]


def wordReport(detail=False, budget=3.0, app=None, hwnd=None):
	app = app or getApp("word", hwnd)
	doc = _firstDoc(app, "ActiveDocument", "Document")
	if doc is None:
		raise OfficeError(_("Could not read the document. It may be busy, for example with a dialog open."))
	deadline = time.time() + budget
	r = Report("word", _safe(lambda: str(doc.Name), ""))

	pages = _int(lambda: doc.ComputeStatistics(2))
	words = _int(lambda: doc.ComputeStatistics(0))
	r.blank = words == 0

	items = [str(s) for s in (_safe(lambda: list(doc.GetCrossReferenceItems(1))) or [])]
	headings = _headingLevels(items)
	tables = _int(lambda: doc.Tables.Count)
	lists = _int(lambda: doc.Lists.Count)
	links = _int(lambda: doc.Hyperlinks.Count)
	inline = _int(lambda: doc.InlineShapes.Count)
	floating = _int(lambda: doc.Shapes.Count)
	comments = _int(lambda: doc.Comments.Count)
	revisions = _int(lambda: doc.Revisions.Count)
	footnotes = _int(lambda: doc.Footnotes.Count)
	endnotes = _int(lambda: doc.Endnotes.Count)

	phrases = counts([
		(pages, "page", "pages"), (words, "word", "words"), (len(headings), "heading", "headings"),
		(tables, "table", "tables"), (lists, "list", "lists"), (links, "link", "links"),
		(inline + floating, "image or object", "images or objects"),
	])
	if phrases:
		r.add(_("Word document has %s.") % ", ".join(phrases))
	review = counts([
		(comments, "comment", "comments"), (revisions, "tracked change", "tracked changes"),
		(footnotes, "footnote", "footnotes"), (endnotes, "endnote", "endnotes"),
	])
	if review:
		r.add(_("Review: %s.") % ", ".join(review))

	# accessibility checks
	missingAlt = []
	for i in range(1, min(inline, 300) + 1):
		if time.time() > deadline:
			break
		shp = _safe(lambda: doc.InlineShapes(i))
		if shp is not None and _int(lambda: shp.Type, 0) in (3, 4, 12, 15):
			if not str(_safe(lambda: shp.AlternativeText, "x") or "").strip():
				missingAlt.append((shp, "inline"))
	for i in range(1, min(floating, 200) + 1):
		if time.time() > deadline:
			break
		shp = _safe(lambda: doc.Shapes(i))
		if shp is not None and _int(lambda: shp.Type, 0) in (3, 11, 13, 24):
			if not str(_safe(lambda: shp.AlternativeText, "x") or "").strip():
				missingAlt.append((shp, "floating"))

	noHeaderRows = []
	tableInfo = []
	for i in range(1, min(tables, 100) + 1):
		if time.time() > deadline:
			break
		tbl = _safe(lambda: doc.Tables(i))
		if tbl is None:
			continue
		rows = _int(lambda: tbl.Rows.Count)
		cols = _int(lambda: tbl.Columns.Count)
		hasHeader = bool(_safe(lambda: tbl.Rows(1).HeadingFormat, False))
		if not hasHeader:
			noHeaderRows.append(i)
		tableInfo.append((i, rows, cols, hasHeader))

	skipped = any(b[0] > a[0] + 1 for a, b in zip(headings, headings[1:]))
	checks = []
	if missingAlt:
		checks.append(plural(len(missingAlt), "image without alternative text", "images without alternative text"))
	if noHeaderRows:
		checks.append(plural(len(noHeaderRows), "table without a header row", "tables without a header row"))
	if skipped:
		checks.append(_("heading levels skip"))
	if checks:
		r.add(_("Check: %s.") % ", ".join(checks))

	if detail:
		r.section(_("Heading outline"), ["h%d: %s" % (lvl, text or _("empty heading")) for lvl, text in headings])
		r.section(_("Tables"), [
			_("Table %d: %s by %s, %s") % (
				i, plural(rows, "row", "rows"), plural(cols, "column", "columns") if cols else _("mixed columns"),
				_("has a header row") if h else _("no header row"),
			) for i, rows, cols, h in tableInfo
		])
		lines = []
		for shp, where in missingAlt[:20]:
			page = _int(lambda: shp.Anchor.Information(3) if where == "floating" else shp.Range.Information(3), 0)
			lines.append(_("%s image on page %s") % (where, page) if page else _("%s image") % where)
		r.section(_("Images without alternative text"), lines)
		cl = []
		for i in range(1, min(comments, 25) + 1):
			c = _safe(lambda: doc.Comments(i))
			if c is not None:
				author = _safe(lambda: str(c.Author), "")
				text = str(_safe(lambda: c.Range.Text, "") or "").strip().replace("\r", " ")[:150]
				cl.append("%s: %s" % (author, text))
		r.section(_("Comments"), cl)
		kinds = {1: ("insertion", "insertions"), 2: ("deletion", "deletions")}
		tally = {}
		authors = set()
		for i in range(1, min(revisions, 300) + 1):
			if time.time() > deadline:
				break
			rev = _safe(lambda: doc.Revisions(i))
			if rev is None:
				continue
			key = kinds.get(_int(lambda: rev.Type, 0), ("other change", "other changes"))
			tally[key] = tally.get(key, 0) + 1
			a = _safe(lambda: str(rev.Author))
			if a:
				authors.add(a)
		lines = [plural(n, k[0], k[1]) for k, n in tally.items()]
		if authors:
			lines.append(_("Authors: %s") % ", ".join(sorted(authors)))
		r.section(_("Tracked changes"), lines)
	return r


# ------------------------------------------------------------------ Excel

def excelReport(detail=False, budget=3.0, app=None, hwnd=None):
	app = app or getApp("excel", hwnd)
	wb = _firstDoc(app, "ActiveWorkbook", "Workbook")
	if wb is None:
		raise OfficeError(_("Could not read the workbook. It may be busy, for example with a dialog open."))
	r = Report("excel", _safe(lambda: str(wb.Name), ""))
	sheets = _int(lambda: wb.Sheets.Count)
	hidden = 0
	for i in range(1, min(sheets, 200) + 1):
		if _int(lambda: wb.Sheets(i).Visible, -1) != -1:
			hidden += 1
	names = _int(lambda: wb.Names.Count)
	text = plural(sheets, "sheet", "sheets")
	if hidden:
		text += _(", %d hidden") % hidden
	r.add(_("Workbook has %s.") % text)

	ws = _safe(lambda: wb.ActiveSheet)
	if ws is not None:
		wsName = _safe(lambda: str(ws.Name), "")
		rows = _int(lambda: ws.UsedRange.Rows.Count)
		cols = _int(lambda: ws.UsedRange.Columns.Count)
		address = _safe(lambda: str(ws.UsedRange.Address), "")
		empty = rows <= 1 and cols <= 1 and not _safe(lambda: ws.UsedRange.Cells(1, 1).Value)
		tables = _int(lambda: ws.ListObjects.Count)
		charts = _int(lambda: ws.ChartObjects().Count)
		shapes = _int(lambda: ws.Shapes.Count)
		comments = _int(lambda: ws.Comments.Count) + _int(lambda: ws.CommentsThreaded.Count)
		links = _int(lambda: ws.Hyperlinks.Count)
		pivots = _int(lambda: ws.PivotTables().Count)
		formulas = _int(lambda: ws.UsedRange.SpecialCells(-4123).Count)
		parts = [] if empty else counts([(rows, "row", "rows"), (cols, "column", "columns")])
		parts += counts([
			(tables, "table", "tables"), (charts, "chart", "charts"), (pivots, "pivot table", "pivot tables"),
			(max(shapes - charts, 0), "other object", "other objects"), (comments, "comment", "comments"),
			(links, "link", "links"), (formulas, "formula", "formulas"),
		])
		if empty:
			r.add(_("Current sheet %s is empty.") % wsName)
		elif parts:
			r.add(_("Current sheet %s has %s.") % (wsName, ", ".join(parts)))
		else:
			r.add(_("Current sheet is %s.") % wsName)
		if _safe(lambda: ws.ProtectContents, False):
			r.add(_("The sheet is protected."))
		r.blank = empty and sheets <= 1

		if detail:
			lines = []
			for i in range(1, min(sheets, 100) + 1):
				sh = _safe(lambda: wb.Sheets(i))
				if sh is None:
					continue
				state = "" if _int(lambda: sh.Visible, -1) == -1 else _(" (hidden)")
				addr = _safe(lambda: str(sh.UsedRange.Address), "")
				lines.append("%s%s %s" % (_safe(lambda: str(sh.Name), ""), state, addr))
			r.section(_("Sheets"), lines)
			lines = []
			for i in range(1, min(tables, 50) + 1):
				lo = _safe(lambda: ws.ListObjects(i))
				if lo is not None:
					lines.append("%s %s" % (_safe(lambda: str(lo.Name), ""), _safe(lambda: str(lo.Range.Address), "")))
			r.section(_("Tables on this sheet"), lines)
			lines = []
			for i in range(1, min(charts, 50) + 1):
				co = _safe(lambda: ws.ChartObjects(i))
				if co is not None:
					lines.append(_safe(lambda: str(co.Name), ""))
			r.section(_("Charts on this sheet"), lines)
			lines = []
			for i in range(1, min(names, 50) + 1):
				n = _safe(lambda: wb.Names(i))
				if n is not None:
					lines.append("%s %s" % (_safe(lambda: str(n.Name), ""), _safe(lambda: str(n.RefersTo), "")))
			r.section(_("Named ranges"), lines)
	return r


# ------------------------------------------------------------------ PowerPoint

def powerpointReport(detail=False, budget=3.0, app=None, hwnd=None):
	app = app or getApp("powerpoint", hwnd)
	pres = _firstDoc(app, "ActivePresentation", "Presentation")
	if pres is None:
		raise OfficeError(_("Could not read the presentation. It may be busy, for example with a dialog open."))
	deadline = time.time() + budget
	r = Report("powerpoint", _safe(lambda: str(pres.Name), ""))
	total = _int(lambda: pres.Slides.Count)
	current = _int(lambda: app.ActiveWindow.View.Slide.SlideIndex)
	r.blank = total <= 1 and _int(lambda: pres.Slides(1).Shapes.Count) == 0

	noTitle = []
	hidden = []
	noAlt = 0
	tables = charts = media = 0
	titles = {}
	scanned = 0
	truncated = False
	for i in range(1, min(total, 300) + 1):
		if time.time() > deadline:
			truncated = True
			break
		slide = _safe(lambda: pres.Slides(i))
		if slide is None:
			continue
		scanned = i
		hasTitle = bool(_safe(lambda: slide.Shapes.HasTitle, False))
		title = ""
		if hasTitle:
			title = str(_safe(lambda: slide.Shapes.Title.TextFrame.TextRange.Text, "") or "").strip()
		titles[i] = title
		if not title:
			noTitle.append(i)
		if _safe(lambda: slide.SlideShowTransition.Hidden, False):
			hidden.append(i)
		for j in range(1, _int(lambda: slide.Shapes.Count) + 1):
			shp = _safe(lambda: slide.Shapes(j))
			if shp is None:
				continue
			stype = _int(lambda: shp.Type, 0)
			if stype in (13, 11):  # picture, linked picture
				if not str(_safe(lambda: shp.AlternativeText, "x") or "").strip():
					noAlt += 1
			if _safe(lambda: shp.HasTable, False):
				tables += 1
			if _safe(lambda: shp.HasChart, False):
				charts += 1
			if stype in (16, 17):  # media, OLE-less movie/sound
				media += 1

	r.add(_("Presentation has %s.") % plural(total, "slide", "slides"))
	if current:
		t = titles.get(current)
		if t is None:
			t = str(_safe(lambda: pres.Slides(current).Shapes.Title.TextFrame.TextRange.Text, "") or "").strip()
		r.add(_("Current slide is %d, %s.") % (current, t or _("no title")))
	phrases = counts([
		(len(hidden), "hidden slide", "hidden slides"), (tables, "table", "tables"),
		(charts, "chart", "charts"), (media, "media clip", "media clips"),
	])
	if phrases:
		r.add(", ".join(phrases).capitalize() + ".")
	checks = []
	if noTitle:
		checks.append(plural(len(noTitle), "slide without a title", "slides without a title", truncated))
	if noAlt:
		checks.append(plural(noAlt, "picture without alternative text", "pictures without alternative text", truncated))
	if checks:
		r.add(_("Check: %s.") % ", ".join(checks))

	if detail:
		r.section(_("Slides"), [
			"%d: %s%s" % (i, titles[i] or _("no title"), _(" (hidden)") if i in hidden else "")
			for i in sorted(titles)
		])
		r.section(_("Slides without a title"), [str(i) for i in noTitle])
	return r


BUILDERS = {"word": wordReport, "excel": excelReport, "powerpoint": powerpointReport}


def buildReport(kind, detail=False, budget=3.0, hwnd=None):
	return BUILDERS[kind](detail=detail, budget=budget, hwnd=hwnd)


def diagnose(kind, hwnd):
	"""Lines describing how (and whether) we can reach the Office application."""
	trace = []
	try:
		app = getApp(kind, hwnd, trace)
		for label, fn in (
			("Application", lambda: app.Name), ("Version", lambda: app.Version), ("Build", lambda: app.Build),
		):
			trace.append("%s: %s" % (label, _safe(fn, "unavailable")))
		item = {"word": lambda: app.ActiveDocument.Name, "excel": lambda: app.ActiveWorkbook.Name,
			"powerpoint": lambda: app.ActivePresentation.Name}[kind]
		trace.append("Active document: %s" % _safe(item, "none (or in Protected View)"))
	except OfficeError as e:
		trace.append("Result: could not connect. " + str(e))
	return trace
