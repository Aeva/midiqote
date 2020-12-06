import time
import string
import rtmidi # via https://github.com/SpotlightKid/python-rtmidi.git
import win32com.client # via package pywin32
import wx # via package wxPython


shell = win32com.client.Dispatch("WScript.Shell")


device_root = 48
device_span = 25

all_symbols = string.ascii_lowercase + string.digits + "."
bass = all_symbols[:device_span]
trebble = all_symbols[12:][:device_span]


NOTE_ON = 0x9
NOTE_OFF = 0x8
POLYPHONIC_PRESSURE = 0xA
CONTROL_CHANGE = 0xB
PROGRAM_CHANGE = 0xC
CHANNEL_PRESSURE = 0xD
PITCH_BEND = 0xE
SYSETM = 0xF


class bard_on:
	def __init__(self):
		self.use_trebble = True
		self.current_port = -1
		self.midiin = rtmidi.MidiIn()
		self.midiin.set_callback(self.midi_event)

	def midi_event(self, event, unused):
		packet, delta = event
		status = packet[0]
		message = status >> 4
		channel = status & 0xF
		if message is NOTE_ON:
			notemap = trebble if self.use_trebble else bass
			note = packet[1]
			symbol = notemap[note-device_root]
			shell.SendKeys(symbol, 0)
		elif message is SYSETM:
			self.use_trebble = (channel & 4) == 0

	def device_changed(self, event):
		old_port = self.current_port
		self.current_port = event.GetSelection()
		if self.current_port != old_port:
			if old_port != -1:
				self.midiin.close_port()
			try:
				self.midiin.open_port(self.current_port)
			except rtmidi._rtmidi.SystemError:
				self.current_port = -1
				event.GetEventObject().SetSelection(-1)
				wx.MessageBox("Bard Error:\nUnable to open MIDI input :(", "Bard-On!", wx.OK | wx.ICON_ERROR)

	def shutdown(self):
		if self.current_port != -1:
			self.midiin.close_port()


class fancy_panel(wx.Panel):
	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		self.SetBackgroundStyle(wx.BG_STYLE_PAINT) 
		self.Bind(wx.EVT_PAINT, self.repaint)
		self.bg_img = wx.Bitmap("window_bg.png")
	
	def repaint(self, event):
		dc = wx.PaintDC(self)
		dc.DrawBitmap(self.bg_img, 0, 0)

if __name__ == "__main__":
	bard = bard_on()
	midi_inputs = bard.midiin.get_ports()

	if len(midi_inputs) == 0:
		app = wx.App()
		wx.MessageBox("Fatal Bard Error:\nNo MIDI input devices found.", "Bard-On!", wx.OK | wx.ICON_ERROR)
		app.MainLoop()
		exit()

	app = wx.App()
	frame = wx.Frame(None, title="Bard-On!")
	frame.SetMinSize(wx.Size(640, 480))
	frame.SetMaxSize(wx.Size(640, 480))
	panel = fancy_panel(frame)

	device_picker = wx.ComboBox(panel, choices=midi_inputs, style=wx.CB_DROPDOWN|wx.CB_READONLY)
	device_picker.SetPosition((400, 100))
	device_picker.SetSize((200, device_picker.GetSize()[1]))
	device_picker.Bind(wx.EVT_COMBOBOX, bard.device_changed)

	#bg_image = wx.StaticBitmap(panel, wx.ID_ANY, wx.Bitmap("windowbg.png", wx.BITMAP_TYPE_ANY))
	#bg_image.SetPosition((0, 0))

	frame.Show()
	app.MainLoop()
	bard.shutdown()
