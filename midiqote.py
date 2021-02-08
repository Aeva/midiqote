import time
import string
from math import floor, ceil
from threading import Thread, Lock, Event
import pygame.midi as midi
import win32api, win32con, win32process # via package pywin32
import wx # via package wxPython


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
#SYMBOLS = string.ascii_lowercase + string.digits + "."
SYMBOLS = list(range(0x41, 0x5B)) + list(range(0x30, 0x3A)) + [0xBE]
F1_KEY = 0x70


def bend_symbol(bend_value):
	if bend_value > 0:
		select = 4 - ceil(bend_value * 8)
		return F1_KEY + select
	elif bend_value < 0:
		select = abs(floor(bend_value * 8)) + 3
		return F1_KEY + select
	else:
		return None


class midiqote(Thread):
	def __init__(self, input_devices):
		Thread.__init__(self, name="midi", daemon=True)
		self.live = True
		self.current_device = None
		self.pending_device = None
		self.device_lock = Lock()
		self.device_changed = Event()
		self.device_changed.clear()

		self.input_devices = input_devices
		self.use_trebble = True
		self.use_ctrl_octave = False
		self.use_prog_octave = False
		self.use_rock_octave = False
		self.transpose = 0
		self.octave = 0
		self.period = 12

		self.last_bend = None
		self.last_mod = None
		self.rest = None

	def run(self):
		self.device_changed.wait()
		while self.live:
			self.device_changed.clear()

			# open the pending midi device
			self.device_lock.acquire()
			assert(self.current_device is None)
			assert(self.pending_device is not None)
			if midi.get_device_info(self.pending_device)[-1] == 1:
				# HACK: portmidi thinks the device is still open, so, uh
				# lets hope the user didn't unplug anything or plug anything new in D:
				midi.quit()
				midi.init()
			self.current_device = midi.Input(self.pending_device)
			self.pending_device = None
			self.device_lock.release()

			# listen for midi events until a new midi device is set
			while self.live and not self.device_changed.is_set():
				while self.live and self.current_device.poll():
					packet, timestamp = self.current_device.read(1)[0]
					status, data1, data2, data3 = packet

					if self.last_mod is not None:
						win32api.keybd_event(self.last_mod, 0, 2, 0)
						self.last_mod = None

					message = status >> 4
					channel = status & 0xF
					if message is NOTE_ON or message is NOTE_OFF:
						note = data1 + self.transpose + self.octave
						while note < ROOT_NOTE:
							note += self.period
						while note > ROOT_NOTE + 36:
							note -= self.period
						note = max(note, ROOT_NOTE)
						note = min(note, ROOT_NOTE + 36)
						symbol = SYMBOLS[note-ROOT_NOTE]
						if message is NOTE_ON:
							win32api.keybd_event(symbol, 0, 0, 0)
							time.sleep(0.03)
						else:
							win32api.keybd_event(symbol, 0, 2, 0)

					elif message is CONTROL_CHANGE:
						control = data1
						value = data2 / 127.0
						if control == 1:
							self.rest_selection = 7 - round(value * 7)
							self.last_mod = self.rest_selection + F1_KEY
							win32api.keybd_event(self.last_mod, 0, 0, 0)

					elif message is PITCH_BEND:
						bend = (((data2 << 7) | data1) / (2**14)) - 0.5
						symbol = bend_symbol(bend)
						if self.last_bend != symbol:
							if self.last_bend is not None:
								win32api.keybd_event(self.last_bend, 0, 2, 0)
							self.last_bend = symbol
						if symbol is not None:
							win32api.keybd_event(symbol, 0, 0, 0)
						else:
							self.last_mod = self.rest_selection + F1_KEY
							win32api.keybd_event(self.last_mod, 0, 0, 0)

					elif self.use_rock_octave and message is SYSETM:
						self.octave = 0 if (channel & 4) == 4 else 12

				time.sleep(0.001)

			# close the current midi device
			self.current_device.close()
			self.current_device = None

	def shutdown(self):
		self.live = False
		self.join()
		midi.quit()

	def set_device(self, event):
		device_id = event if type(event) is int else self.input_devices[event.GetInt()][0]
		self.device_lock.acquire()
		self.pending_device = device_id
		self.device_lock.release()
		self.device_changed.set()

	def set_ctrl_octave(self, event):
		self.use_ctrl_octave = event.IsChecked()

	def set_prog_octave(self, event):
		self.use_prog_octave = event.IsChecked()

	def set_rock_octave(self, event):
		self.use_rock_octave = event.IsChecked()

	def set_transpose(self, event):
		self.transpose = MIDDLE_C - event.GetInt()

	def set_period(self, event):
		self.period = event.GetInt()
		assert(self.period > 0)
		assert(self.period < len(SYMBOLS))


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
	pid = win32api.GetCurrentProcessId()
	handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, True, pid)
	win32process.SetPriorityClass(handle, win32process.REALTIME_PRIORITY_CLASS)

	midi.init()
	all_devices = [midi.get_device_info(i) for i in range(midi.get_count())]
	input_devices = [(i, d[1]) for (i, d) in enumerate(all_devices) if d[2] == 1]
	input_names = [name for (i, name) in input_devices]
	default_device = midi.get_default_input_id()
	try:
		default_selection = [i for (i, name) in input_devices].index(default_device)
	except ValueError:
		default_selection = None

	midi_thread = midiqote(input_devices)
	midi_thread.start()
	midi_thread.set_device(default_device)

	if len(input_names) == 0:
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

	device_picker = wx.ComboBox(panel, choices=input_names, style=wx.CB_DROPDOWN|wx.CB_READONLY)
	device_picker.Bind(wx.EVT_COMBOBOX, midi_thread.set_device)
	device_picker.SetSelection(default_selection)
	vbox.Add(device_picker, border=100, flag=wx.TOP)

	#midi_ctrl_checkbox = wx.CheckBox(panel, label="All MIDI control toggles\nas octave up/down.", style=wx.CHK_2STATE)
	#midi_ctrl_checkbox.Bind(wx.EVT_CHECKBOX, midi_thread.set_ctrl_octave)
	#vbox.Add(midi_ctrl_checkbox, border=10, flag=wx.TOP)

	#midi_prog_checkbox = wx.CheckBox(panel, label="All MIDI program toggles\nas octave up/down.", style=wx.CHK_2STATE)
	#midi_prog_checkbox.Bind(wx.EVT_CHECKBOX, midi_thread.set_prog_octave)
	#vbox.Add(midi_prog_checkbox, border=10, flag=wx.TOP)

	rockband_checkbox = wx.CheckBox(panel, label="Rockband MIDI controller\nstart/select as low/high octave.", style=wx.CHK_2STATE)
	rockband_checkbox.Bind(wx.EVT_CHECKBOX, midi_thread.set_rock_octave)
	vbox.Add(rockband_checkbox, border=10, flag=wx.TOP)

	middle_c_hbox = wx.BoxSizer(wx.HORIZONTAL)
	vbox.Add(middle_c_hbox, border=10, flag=wx.TOP)
	middle_c = wx.SpinCtrl(panel, min=1, max=127, initial=60)
	middle_c.Bind(wx.EVT_SPINCTRL, midi_thread.set_transpose)
	middle_c_hbox.Add(middle_c)
	middle_c_label = wx.StaticText(panel, style=wx.ALIGN_LEFT, label="Middle C")
	middle_c_hbox.Add(middle_c_label, flag=wx.ALIGN_CENTER_VERTICAL)

	period_hbox = wx.BoxSizer(wx.HORIZONTAL)
	vbox.Add(period_hbox, border=10, flag=wx.TOP)
	period = wx.SpinCtrl(panel, min=1, max=36, initial=12)
	period.Bind(wx.EVT_SPINCTRL, midi_thread.set_period)
	period_hbox.Add(period)
	period_label = wx.StaticText(panel, style=wx.ALIGN_LEFT, label="Boundary Period")
	period_hbox.Add(period_label, flag=wx.ALIGN_CENTER_VERTICAL)

	panel.SetSizer(hbox)
	frame.Show()
	app.MainLoop()
	midi_thread.shutdown()

