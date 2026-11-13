# Page Summary - reports for browse mode documents: web pages, PDFs and e-books.
import time

import addonHandler
import api

from .report import Report, counts

addonHandler.initTranslation()

# (quick nav type, singular, plural)
COUNT_TYPES = [
	("landmark", "landmark", "landmarks"),
	("heading", "heading", "headings"),
	("link", "link", "links"),
	("formField", "form field", "form fields"),
	("table", "table", "tables"),
	("list", "list", "lists"),
	("graphic", "graphic", "graphics"),
]

LEADS = {
	"web": "Page has %s.",
	"pdf": "PDF has %s.",
	"ebook": "Book page has %s.",
	"document": "Document has %s.",
}


def scan(ti, itemType, withLabels=False, positions=False, deadline=None):
	"""Walk all nodes of a quick nav type.

	Returns (list of (position, label), truncated), or (None, False) if the type is unsupported.
	"""
	results = []
	truncated = False
	try:
		for item in ti._iterNodesByType(itemType):
			if deadline is not None and time.time() > deadline:
				truncated = True
				break
			pos = len(results)
			if positions:
				try:
					pos = item.textInfo._startOffset
				except Exception:
					pass
			label = ""
			if withLabels:
				try:
					label = item.label or ""
				except Exception:
					label = ""
			results.append((pos, label))
	except Exception:
		return None, False
	return results, truncated


def pageTitle(ti):
	for getter in (lambda: ti.rootNVDAObject.name, lambda: api.getForegroundObject().name):
		try:
			name = getter()
			if name:
				return name
		except Exception:
			pass
	return ""


def gather(ti, detail=False, budget=4.0):
	"""Counts of elements. With detail=True also the heading outline and landmark names."""
	deadline = time.time() + budget
	data = {"counts": {}, "truncated": set(), "headings": [], "landmarks": []}

	if detail:
		headings = []
		for level in range(1, 7):
			items, cut = scan(ti, "heading%d" % level, True, True, deadline)
			if items is None:
				headings = None
				break
			if cut:
				data["truncated"].add("heading")
			headings.extend((pos, level, label) for pos, label in items)
		if headings is None:
			items, cut = scan(ti, "heading", True, True, deadline)
			headings = [(p, 0, l) for p, l in (items or [])]
			if cut:
				data["truncated"].add("heading")
		headings.sort(key=lambda h: h[0])
		data["headings"] = headings
		data["counts"]["heading"] = len(headings)

	for itemType, _s, _p in COUNT_TYPES:
		if detail and itemType == "heading":
			continue
		wantLabels = detail and itemType == "landmark"
		items, cut = scan(ti, itemType, wantLabels, wantLabels, deadline)
		if items is None:
			continue
		data["counts"][itemType] = len(items)
		if cut:
			data["truncated"].add(itemType)
		if wantLabels:
			data["landmarks"] = items
	return data


def contentLength(ti):
	"""Length of the document text, or None if unknown. Used to tell when a PDF has finished loading."""
	try:
		import textInfos
		return ti.makeTextInfo(textInfos.POSITION_LAST)._endOffset
	except Exception:
		return None


def buildReport(ti, kind="web", detail=False, budget=4.0):
	data = gather(ti, detail, budget)
	r = Report(kind, pageTitle(ti))
	c = data["counts"]
	phrases = counts([
		(c.get(t, 0), s, p, t in data["truncated"]) for t, s, p in COUNT_TYPES
	])
	if phrases:
		r.add(_(LEADS.get(kind, LEADS["document"])) % ", ".join(phrases))

	structure = sum(c.get(t, 0) for t in ("heading", "list", "table", "graphic", "landmark"))
	if kind == "pdf" and c and not structure and not data["truncated"]:
		r.add(_(
			"No headings, lists or tables were found. This PDF may be untagged or scanned, "
			"so reading order and structure may be poor."
		))

	if detail:
		lines = []
		for t, s, p in COUNT_TYPES:
			n = c.get(t)
			if n is not None:
				lines.extend(counts([(n, s, p, t in data["truncated"])]) or [_("no %s") % p])
		r.section(_("Totals"), lines)
		r.section(
			_("Landmarks"), [label or _("unnamed landmark") for _pos, label in data["landmarks"]]
		)
		outline = []
		prev = 0
		skipped = False
		for _pos, level, label in data["headings"]:
			outline.append("%s%s" % (("h%d: " % level) if level else "", label or _("empty heading")))
			if level and prev and level > prev + 1:
				skipped = True
			if level:
				prev = level
		r.section(_("Heading outline"), outline)
		notes = []
		if skipped:
			notes.append(_("Heading levels skip, for example from h2 to h4."))
		if kind == "pdf" and c and not structure:
			notes.append(_("No document structure found; the PDF may be untagged."))
		r.section(_("Accessibility notes"), notes)
	return r
