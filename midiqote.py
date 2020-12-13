import time
import string
import ctypes
import ctypes.wintypes
import win32com.client # via package pywin32
import wx # via package wxPython


shell = win32com.client.Dispatch("WScript.Shell")

winmm = ctypes.CDLL("winmm.dll")
winmm.midiInOpen.restype = ctypes.c_uint
winmm.midiInGetNumDevs.restype = ctypes.c_uint
winmm.midiInStart.restype = ctypes.c_uint
winmm.midiInStop.restype = ctypes.c_uint
winmm.midiInGetErrorTextW.restype = ctypes.c_uint


def result(error_code):
	message = ctypes.create_string_buffer(1024)
	result = winmm.midiInGetErrorTextW(error_code, message, 1024)
	return message.value if result else None


def device_name(device):
	class capabilities(ctypes.Structure):
		_fields_ = [
			("wMid", ctypes.wintypes.WORD),
			("wPid", ctypes.wintypes.WORD),
			("vDriverVersion", ctypes.wintypes.UINT),
			("szPname", ctypes.wintypes.WCHAR * 32),
			("dwSupport", ctypes.wintypes.DWORD)]
	info = capabilities()
	device = ctypes.wintypes.UINT(device)
	size = ctypes.wintypes.UINT(info.__sizeof__())
	error = result(winmm.midiInGetDevCapsW(device, ctypes.byref(info), size))
	if error:
		raise(error)
	return info.szPname


def midi_devices():
	device_count = winmm.midiInGetNumDevs()
	return [device_name(i) for i in range(device_count)]


def midi_in_open(port, callback):
	handle = ctypes.c_uint()
	midi_in_callback = ctypes.CFUNCTYPE(None, ctypes.c_uint, ctypes.c_uint, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint)
	print("midi open")
	error = result(winmm.midiInOpen(ctypes.byref(handle), port, midi_in_callback(callback), ctypes.c_void_p(), 0x00030000))
	return handle, error


def midi_in_start(handle):
	print("midi start")
	return result(winmm.midiInStart(handle))


def midi_in_stop(handle):
	return result(winmm.midiInStop(handle))


NOTE_ON = 0x9
NOTE_OFF = 0x8
POLYPHONIC_PRESSURE = 0xA
CONTROL_CHANGE = 0xB
PROGRAM_CHANGE = 0xC
CHANNEL_PRESSURE = 0xD
PITCH_BEND = 0xE
SYSETM = 0xF
MIDDLE_C = 60
ROOT_NOTE = 48
SYMBOLS = string.ascii_lowercase + string.digits + "."


class midiqote:
	def __init__(self):
		self.use_trebble = True
		self.current_port = -1
		self.use_ctrl_octave = False
		self.use_prog_octave = False
		self.use_rock_octave = False
		self.transpose = 0
		self.octave = 0
		self.handle = None

	def midi_event(self, handle, msg, custom, status, data):
		print(hex(msg), status, data)
		message = status >> 4
		channel = status & 0xF		
		if message is NOTE_ON:
			note = data + self.transpose + self.octave
			if note >= ROOT_NOTE and note <= (ROOT_NOTE + 37):
				symbol = SYMBOLS[note-ROOT_NOTE]
				shell.SendKeys(symbol, 0)
		elif self.use_rock_octave and message is SYSETM:
			self.octave = 0 if (channel & 4) == 4 else 12

	def device_changed(self, event, override = None):
		old_port = self.current_port
		self.current_port = override if override is not None else event.GetSelection()
		if self.current_port != old_port:
			if old_port != -1:
				midi_in_stop(self.handle)
				self.handle = None
			self.handle, error = midi_in_open(self.current_port, self.midi_event)
			if (error):
				self.current_port = -1
				event.GetEventObject().SetSelection(-1)
				wx.MessageBox("Bard Error:\nUnable to open MIDI input :(\n" + error, "Midiqo'te", wx.OK | wx.ICON_ERROR)
				return
			error = midi_in_start(self.handle)
			if (error):
				self.current_port = -1
				event.GetEventObject().SetSelection(-1)
				wx.MessageBox("Bard Error:\nUnable to start MIDI input :(\n" + error, "Midiqo'te", wx.OK | wx.ICON_ERROR)
				return

	def set_ctrl_octave(self, event):
		self.use_ctrl_octave = event.IsChecked()

	def set_prog_octave(self, event):
		self.use_prog_octave = event.IsChecked()

	def set_rock_octave(self, event):
		self.use_rock_octave = event.IsChecked()

	def set_transpose(self, event):
		self.transpose = MIDDLE_C - event.GetInt()

	def shutdown(self):
		if self.current_port != -1:
			midi_in_stop(self.handle)
			self.handle = None


class fancy_panel(wx.Panel):
	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		self.SetBackgroundStyle(wx.BG_STYLE_PAINT) 
		self.Bind(wx.EVT_PAINT, self.repaint)
		self.bg_img = wx.Bitmap("window_bg.png")
	
	def repaint(self, event):
		dc = wx.PaintDC(self)
		dc.DrawBitmap(self.bg_img, 0, -39)


if __name__ == "__main__":
	bard = midiqote()
	midi_inputs = midi_devices()

	if len(midi_inputs) == 0:
		app = wx.App()
		wx.MessageBox("Fatal Bard Error:\nNo MIDI input devices found.", "Midiqo'te", wx.OK | wx.ICON_ERROR)
		app.MainLoop()
		exit()

	app = wx.App()
	frame = wx.Frame(None, title="Midiqo'te")
	frame.SetMinSize(wx.Size(640, 480))
	frame.SetMaxSize(wx.Size(640, 480))
	panel = fancy_panel(frame)
	hbox = wx.BoxSizer(wx.HORIZONTAL)
	vbox = wx.BoxSizer(wx.VERTICAL)
	hbox.Add(vbox, border=400, flag=wx.LEFT)
	
	device_picker = wx.ComboBox(panel, choices=midi_inputs, style=wx.CB_DROPDOWN|wx.CB_READONLY)
	device_picker.Bind(wx.EVT_COMBOBOX, bard.device_changed)
	device_picker.SetSelection(0)
	bard.device_changed(None, 0)
	vbox.Add(device_picker, border=100, flag=wx.TOP)

	#midi_ctrl_checkbox = wx.CheckBox(panel, label="All MIDI control toggles\nas octave up/down.", style=wx.CHK_2STATE)
	#midi_ctrl_checkbox.Bind(wx.EVT_CHECKBOX, bard.set_ctrl_octave)
	#vbox.Add(midi_ctrl_checkbox, border=10, flag=wx.TOP)

	#midi_prog_checkbox = wx.CheckBox(panel, label="All MIDI program toggles\nas octave up/down.", style=wx.CHK_2STATE)
	#midi_prog_checkbox.Bind(wx.EVT_CHECKBOX, bard.set_prog_octave)
	#vbox.Add(midi_prog_checkbox, border=10, flag=wx.TOP)

	rockband_checkbox = wx.CheckBox(panel, label="Rockband MIDI controller\nstart/select as low/high octave.", style=wx.CHK_2STATE)
	rockband_checkbox.Bind(wx.EVT_CHECKBOX, bard.set_rock_octave)
	vbox.Add(rockband_checkbox, border=10, flag=wx.TOP)

	middle_c_hbox = wx.BoxSizer(wx.HORIZONTAL)
	vbox.Add(middle_c_hbox, border=10, flag=wx.TOP)
	middle_c = wx.SpinCtrl(panel, min=1, max=127, initial=60)
	middle_c.Bind(wx.EVT_SPINCTRL, bard.set_transpose)
	middle_c_hbox.Add(middle_c)

	middle_c_label = wx.StaticText(panel, style=wx.ALIGN_LEFT, label="Middle C")
	middle_c_hbox.Add(middle_c_label, flag=wx.ALIGN_CENTER_VERTICAL)

	panel.SetSizer(hbox)
	frame.Show()
	app.MainLoop()
	bard.shutdown()
