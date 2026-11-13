# Page Summary - settings panel (NVDA menu > Preferences > Settings > Page Summary).
import addonHandler
import config
import wx
from gui import guiHelper
from gui.settingsDialogs import SettingsPanel

addonHandler.initTranslation()

CONFIG_SPEC = {
	"autoAnnounce": "boolean(default=True)",
	"autoWeb": "boolean(default=True)",
	"autoPdf": "boolean(default=True)",
	"autoEbook": "boolean(default=True)",
	"autoOffice": "boolean(default=True)",
	"autoDelay": "integer(default=1000, min=0, max=5000)",
	"extraEbookApps": "string(default='')",
}


class PageSummarySettingsPanel(SettingsPanel):
	title = _("Page Summary")

	def makeSettings(self, settingsSizer):
		cfg = config.conf["pageSummary"]
		helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

		def check(label, key):
			box = helper.addItem(wx.CheckBox(self, label=label))
			box.SetValue(cfg[key])
			return box

		self.auto = check(_("&Announce a summary automatically when a document loads"), "autoAnnounce")
		self.web = check(_("For &web pages"), "autoWeb")
		self.pdf = check(_("For &PDF documents"), "autoPdf")
		self.ebook = check(_("For &e-books"), "autoEbook")
		self.office = check(_("For Microsoft &Office documents (Word, Excel, PowerPoint)"), "autoOffice")
		self.delay = helper.addLabeledControl(
			_("&Delay after the document loads, in milliseconds:"),
			wx.SpinCtrl, min=0, max=5000, initial=cfg["autoDelay"],
		)
		self.extra = helper.addLabeledControl(
			_("Additional e-&book reader program names (comma separated, for example thorium, ebook-viewer):"),
			wx.TextCtrl,
		)
		self.extra.SetValue(cfg["extraEbookApps"])

	def onSave(self):
		cfg = config.conf["pageSummary"]
		cfg["autoAnnounce"] = self.auto.GetValue()
		cfg["autoWeb"] = self.web.GetValue()
		cfg["autoPdf"] = self.pdf.GetValue()
		cfg["autoEbook"] = self.ebook.GetValue()
		cfg["autoOffice"] = self.office.GetValue()
		cfg["autoDelay"] = self.delay.GetValue()
		cfg["extraEbookApps"] = self.extra.GetValue().strip()
