# Page Summary - JAWS-style summary of web pages, PDFs, e-books and Microsoft Office documents.
import ctypes
import struct

import addonHandler
import api
import browseMode
import config
import globalPluginHandler
import gui
import scriptHandler
import ui
import wx
from logHandler import log
from scriptHandler import script

from . import office, webdoc
from .settings import CONFIG_SPEC, PageSummarySettingsPanel

try:
	import virtualBuffers
except ImportError:  # pragma: no cover
	virtualBuffers = None

addonHandler.initTranslation()

# Seconds allowed for scanning.
AUTO_BUDGET = 2.0
SUMMARY_BUDGET = 4.0
DETAIL_BUDGET = 12.0

BROWSER_APPS = {
	"chrome", "msedge", "firefox", "brave", "opera", "vivaldi", "chromium", "iexplore",
	"waterfox", "librewolf", "arc",
}
PDF_APPS = {"acrord32", "acrord64", "acrobat", "acrord32info"}
EBOOK_APPS = {"thorium", "thoriumreader", "ebook-viewer", "calibre", "kindle", "koodoreader", "readium"}
# app module name -> (main window class, document kind)
OFFICE_APPS = {
	"winword": ("OpusApp", "word"),
	"excel": ("XLMAIN", "excel"),
	"powerpnt": ("PPTFrameClass", "powerpoint"),
}
AUTO_KEYS = {
	"web": "autoWeb", "pdf": "autoPdf", "ebook": "autoEbook",
	"word": "autoOffice", "excel": "autoOffice", "powerpoint": "autoOffice",
}


def _getDocument():
	"""The active browse mode document, or None."""
	try:
		ti = api.getFocusObject().treeInterceptor
	except Exception:
		return None
	if ti is None or not isinstance(ti, browseMode.BrowseModeDocumentTreeInterceptor):
		return None
	return ti


def _appName(ti):
	try:
		return ti.rootNVDAObject.appModule.appName.lower()
	except Exception:
		return ""


def _docId(ti):
	try:
		return ti.documentConstantIdentifier or ""
	except Exception:
		return ""


def _looksLikePdf(url, title):
	path = url.lower().split("#")[0].split("?")[0]
	if path.endswith(".pdf") or "pdf.js" in path or "mhjfbmdgcfjbbpaeojofohoefgiehjai" in path:
		return True
	return (title or "").lower().endswith(".pdf")


def classify(ti, extraEbookApps=()):
	"""'web', 'pdf' or 'ebook' for browse mode documents we announce automatically, else None."""
	app = _appName(ti)
	if app in PDF_APPS:
		return "pdf"
	if app in EBOOK_APPS or app in extraEbookApps:
		return "ebook"
	if app in BROWSER_APPS:
		url = _docId(ti)
		try:
			title = ti.rootNVDAObject.name
		except Exception:
			title = ""
		if _looksLikePdf(url, title):
			return "pdf"
		if url.lower().split("#")[0].split("?")[0].endswith(".epub"):
			return "ebook"
		return "web"
	return None


def _extraEbookApps():
	raw = config.conf["pageSummary"]["extraEbookApps"]
	return {a.strip().lower() for a in raw.split(",") if a.strip()}


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("Page Summary")

	def __init__(self):
		super().__init__()
		config.conf.spec["pageSummary"] = CONFIG_SPEC
		gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(PageSummarySettingsPanel)
		self._patches = []
		self._pending = None
		self._officeSeen = set()
		self._installHooks()

	def terminate(self):
		try:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(PageSummarySettingsPanel)
		except ValueError:
			pass
		self._stopPending()
		for cls, name, own, orig in reversed(self._patches):
			try:
				if own:
					setattr(cls, name, orig)
				else:
					delattr(cls, name)
			except Exception:
				log.debugWarning("Page Summary: could not restore %s.%s" % (cls.__name__, name), exc_info=True)
		self._patches = []

	def _stopPending(self):
		if self._pending is not None:
			try:
				self._pending.Stop()
			except Exception:
				pass
			self._pending = None

	# ------------------------------------------------------------ browse mode documents (web, PDF, e-book)

	def _patch(self, cls, name, wrapperFactory):
		orig = getattr(cls, name, None)
		if orig is None:
			return
		self._patches.append((cls, name, name in cls.__dict__, orig))
		setattr(cls, name, wrapperFactory(orig))

	def _installHooks(self):
		plugin = self

		def wrapGain(orig):
			def patched(ti, *args, **kwargs):
				result = orig(ti, *args, **kwargs)
				plugin._onDocumentEvent(ti)
				return result
			return patched

		def wrapLoad(orig):
			def patched(ti, *args, **kwargs):
				result = orig(ti, *args, **kwargs)
				success = kwargs.get("success", args[0] if args else True)
				if success:
					plugin._onDocumentEvent(ti)
				return result
			return patched

		try:
			self._patch(browseMode.BrowseModeDocumentTreeInterceptor, "event_treeInterceptor_gainFocus", wrapGain)
			if virtualBuffers is not None:
				vb = virtualBuffers.VirtualBuffer
				if "event_treeInterceptor_gainFocus" in vb.__dict__:
					self._patch(vb, "event_treeInterceptor_gainFocus", wrapGain)
				if "_loadBufferDone" in vb.__dict__:
					self._patch(vb, "_loadBufferDone", wrapLoad)
		except Exception:
			log.error("Page Summary: could not hook document loading", exc_info=True)

	def _onDocumentEvent(self, ti):
		"""A browse mode document finished loading or gained focus. Never raises."""
		try:
			cfg = config.conf["pageSummary"]
			if not cfg["autoAnnounce"] or not getattr(ti, "isReady", False):
				return
			kind = classify(ti, _extraEbookApps())
			if kind is None or not cfg[AUTO_KEYS[kind]]:
				return
			if api.getFocusObject().treeInterceptor is not ti:
				return  # background tab: announce when it first gets focus
			marker = _docId(ti)
			if hasattr(ti, "_pageSummaryId") and ti._pageSummaryId == marker:
				return  # already announced for this document
			ti._pageSummaryId = marker
			self._schedule(ti, kind)
		except Exception:
			log.debugWarning("Page Summary: auto announce check failed", exc_info=True)

	def _schedule(self, ti, kind):
		sayAllOnLoad = False
		try:
			sayAllOnLoad = bool(config.conf["browseMode"]["autoSayAllOnPageLoad"])
		except Exception:
			pass
		delay = int(config.conf["pageSummary"]["autoDelay"])
		if kind == "pdf":
			delay = max(delay, 2000)  # PDF content is added to the document after it loads
		elif sayAllOnLoad or delay <= 0:
			self._announceAuto(ti, kind)  # speak before Say All starts
			return
		self._stopPending()
		self._pending = wx.CallLater(delay, self._announceAuto, ti, kind)

	def _announceAuto(self, ti, kind, attempt=0):
		try:
			if not getattr(ti, "isReady", False) or api.getFocusObject().treeInterceptor is not ti:
				return
			if kind == "pdf":
				length = webdoc.contentLength(ti)
				if length is not None and length < 200:
					if attempt < 3:
						self._pending = wx.CallLater(2000, self._announceAuto, ti, kind, attempt + 1)
					return
			text = webdoc.buildReport(ti, kind, False, AUTO_BUDGET).spoken()
			if text:
				ui.message(text)
		except Exception:
			log.debugWarning("Page Summary: auto announce failed", exc_info=True)

	# ------------------------------------------------------------ Microsoft Office documents

	def event_gainFocus(self, obj, nextHandler):
		try:
			self._checkOffice(obj)
		except Exception:
			log.debugWarning("Page Summary: Office check failed", exc_info=True)
		nextHandler()

	def _checkOffice(self, obj):
		cfg = config.conf["pageSummary"]
		if not cfg["autoAnnounce"] or not cfg["autoOffice"]:
			return
		info = OFFICE_APPS.get((obj.appModule.appName or "").lower())
		if not info:
			return
		fg = api.getForegroundObject()
		if fg is None or fg.windowClassName != info[0]:
			return
		key = (info[1], fg.windowHandle, fg.name or "")
		if key in self._officeSeen:
			return
		if len(self._officeSeen) > 500:
			self._officeSeen.clear()
		self._officeSeen.add(key)
		delay = max(int(cfg["autoDelay"]), 1500)  # let NVDA finish announcing the document first
		self._stopPending()
		self._pending = wx.CallLater(delay, self._announceOffice, info[1], key)

	def _announceOffice(self, kind, key):
		try:
			fg = api.getForegroundObject()
			if fg is None or (kind, fg.windowHandle, fg.name or "") != key:
				return  # the user has moved on
			report = office.buildReport(kind, False, 3.0, hwnd=key[1])
			if report.blank or report.empty:
				return
			ui.message(report.spoken())
		except office.OfficeError:
			log.debugWarning("Page Summary: Office document not available", exc_info=True)
		except Exception:
			log.debugWarning("Page Summary: Office announce failed", exc_info=True)

	# ------------------------------------------------------------ scripts

	def _context(self):
		"""(kind, browse mode document or None, window handle) for whatever has focus, or None."""
		try:
			fg = api.getForegroundObject()
			info = OFFICE_APPS.get((fg.appModule.appName or "").lower())
			if info and fg.windowClassName == info[0]:
				return info[1], None, fg.windowHandle
		except Exception:
			pass
		ti = _getDocument()
		if ti is None:
			return None
		return (classify(ti, _extraEbookApps()) or "document"), ti, 0

	@script(
		description=_(
			"Speaks a summary of the current web page, PDF, e-book or Office document. "
			"Press twice quickly for details."
		),
		gesture="kb:NVDA+shift+q",
	)
	def script_pageSummary(self, gesture):
		ctx = self._context()
		if ctx is None:
			ui.message(
				_("Not in a web page, PDF, e-book or Office document. Press NVDA plus Alt plus Shift plus Q for diagnostics.")
			)
			return
		kind, ti, hwnd = ctx
		if ti is not None and not getattr(ti, "isReady", True):
			ui.message(_("Document is still loading."))
			return
		detail = scriptHandler.getLastScriptRepeatCount() >= 1
		if not detail:
			ui.message(_("Summarizing"))
		try:
			if ti is None:
				report = office.buildReport(kind, detail, DETAIL_BUDGET if detail else SUMMARY_BUDGET, hwnd=hwnd)
			else:
				report = webdoc.buildReport(ti, kind, detail, DETAIL_BUDGET if detail else SUMMARY_BUDGET)
		except office.OfficeError as e:
			ui.message(str(e) + " " + _("Press NVDA plus Alt plus Shift plus Q for diagnostics."))
			return
		if detail:
			ui.browseableMessage(report.html(), _("Document summary"), isHtml=True)
		else:
			ui.message(
				report.spoken(withTitle=True, withHint=True)
				or _("Nothing to summarize: no headings, links, landmarks or form fields found.")
			)

	@script(
		description=_("Turns the automatic summary on or off when documents load."),
		gesture="kb:NVDA+control+shift+q",
	)
	def script_toggleAuto(self, gesture):
		cfg = config.conf["pageSummary"]
		cfg["autoAnnounce"] = not cfg["autoAnnounce"]
		ui.message(_("Automatic summary on") if cfg["autoAnnounce"] else _("Automatic summary off"))

	# ------------------------------------------------------------ diagnostics

	def _diagnosticLines(self):
		lines = []

		def add(text):
			lines.append(text)

		try:
			import buildVersion
			add("NVDA version: %s" % buildVersion.version)
		except Exception:
			pass
		try:
			add("Add-on version: %s" % addonHandler.getCodeAddon().version)
		except Exception:
			pass
		add("NVDA process: %d-bit" % (struct.calcsize("P") * 8))
		try:
			add("NVDA running as administrator: %s" % ("yes" if ctypes.windll.shell32.IsUserAnAdmin() else "no"))
		except Exception:
			pass
		try:
			fg = api.getForegroundObject()
			add("Foreground window: program %r, class %r, title %r, handle 0x%x" % (
				fg.appModule.appName, fg.windowClassName, fg.name, fg.windowHandle or 0))
			focus = api.getFocusObject()
			add("Focused object: program %r, class %r, role %s" % (
				focus.appModule.appName, focus.windowClassName, focus.role))
			info = OFFICE_APPS.get((fg.appModule.appName or "").lower())
			if info is None:
				add("This program is not one of Word, Excel or PowerPoint.")
			elif fg.windowClassName != info[0]:
				add("Expected the main window class %r for %s but found %r (a dialog or pane may have focus)." % (
					info[0], info[1], fg.windowClassName))
			else:
				add("Detected an Office %s window. Connecting:" % info[1])
				lines.extend("  " + l for l in office.diagnose(info[1], fg.windowHandle))
		except Exception as e:
			add("Could not read the foreground window: %s" % e)
		ti = _getDocument()
		if ti is None:
			add("Browse mode document: none at the focus.")
		else:
			add("Browse mode document: program %r, address %r, ready: %s, detected type: %s" % (
				_appName(ti), _docId(ti), getattr(ti, "isReady", "unknown"),
				classify(ti, _extraEbookApps()) or "not announced automatically"))
			try:
				add("Counts: %s" % webdoc.gather(ti, False, 2.0)["counts"])
			except Exception as e:
				add("Counting failed: %s" % e)
		try:
			cfg = config.conf["pageSummary"]
			add("Settings: " + ", ".join("%s=%s" % (k, cfg[k]) for k in CONFIG_SPEC))
		except Exception:
			pass
		return lines

	@script(
		description=_("Shows diagnostic information about why a summary may not work in the current document."),
		gesture="kb:NVDA+alt+shift+q",
	)
	def script_diagnostics(self, gesture):
		import html
		lines = self._diagnosticLines()
		log.info("Page Summary diagnostics:\n" + "\n".join(lines))
		ui.browseableMessage(
			"<h1>Page Summary diagnostics</h1><ul>%s</ul>" % "".join("<li>%s</li>" % html.escape(l) for l in lines),
			_("Page Summary diagnostics"), isHtml=True,
		)
