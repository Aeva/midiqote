import time
import string
import rtmidi # via https://github.com/SpotlightKid/python-rtmidi.git
import win32com.client # via package pywin32


shell = win32com.client.Dispatch("WScript.Shell")


device_name = 'USB Midi Cable 1'
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


use_trebble = True
def midi_event(event, unused):
	global use_trebble
	packet, delta = event
	status = packet[0]
	message = status >> 4
	channel = status & 0xF
	if message is NOTE_ON:
		notemap = trebble if use_trebble else bass
		note = packet[1]
		symbol = notemap[note-device_root]
		shell.SendKeys(symbol, 0)
	elif message is SYSETM:
		use_trebble = (channel & 4) == 0


if __name__ == "__main__":
	midiin = rtmidi.MidiIn()
	available_inputs = midiin.get_ports()
	assert(device_name in available_inputs)
	port_number = available_inputs.index(device_name)
	midiin.open_port(port_number)
	midiin.set_callback(midi_event)
	print("Listening to %s" % device_name)
	try:
		while True:
			time.sleep(1.0)
	except KeyboardInterrupt:
		pass
	midiin.close_port()
