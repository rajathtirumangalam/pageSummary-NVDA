# Page Summary - a small report object shared by every document type.
import html as _html

import addonHandler

addonHandler.initTranslation()

KIND_LABELS = {
	"web": "Web page",
	"pdf": "PDF document",
	"ebook": "E-book",
	"document": "Document",
	"word": "Word document",
	"excel": "Excel workbook",
	"powerpoint": "PowerPoint presentation",
}


def plural(n, singular, pluralForm, atLeast=False):
	text = "%s %s" % ("{:,}".format(n), singular if n == 1 else pluralForm)
	return (_("at least %s") % text) if atLeast else text


def counts(items):
	"""items: (n, singular, plural[, atLeast]); zero counts are skipped. Returns a list of phrases."""
	out = []
	for item in items:
		n, s, p = item[:3]
		if n:
			out.append(plural(n, s, p, item[3] if len(item) > 3 else False))
	return out


class Report:
	def __init__(self, kind, title=""):
		self.kind = kind
		self.title = title or ""
		self.sentences = []
		self.sections = []  # (heading, [lines])
		self.blank = False  # nothing worth announcing

	def add(self, sentence):
		if sentence:
			self.sentences.append(sentence)

	def section(self, heading, lines):
		if lines:
			self.sections.append((heading, list(lines)))

	@property
	def empty(self):
		return not self.sentences

	def spoken(self, withTitle=False, withHint=False):
		text = " ".join(self.sentences)
		if not text:
			return ""
		if withTitle and self.title:
			text = self.title + ". " + text
		if withHint:
			text += " " + _("Press twice quickly for details.")
		return text

	def html(self):
		out = ["<h1>%s</h1>" % _html.escape(self.title or _(KIND_LABELS.get(self.kind, "Document")))]
		out.append("<p>%s</p>" % _html.escape(_(KIND_LABELS.get(self.kind, "Document"))))
		for s in self.sentences:
			out.append("<p>%s</p>" % _html.escape(s))
		for heading, lines in self.sections:
			out.append("<h2>%s</h2><ul>" % _html.escape(heading))
			for line in lines:
				out.append("<li>%s</li>" % _html.escape(line))
			out.append("</ul>")
		return "".join(out)
