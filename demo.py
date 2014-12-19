#!/usr/bin/python
# -*- coding: utf-8 -*-
import win32api
import ctypes
from win32con import *


# add ../../Sources to the PYTHONPATH
from yoctopuce.yocto_api import *
from yoctopuce.yocto_anbutton import *

SendInput = ctypes.windll.user32.SendInput

# C struct redefinitions
PUL = ctypes.POINTER(ctypes.c_ulong)


class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                ("mi", MouseInput),
                ("hi", HardwareInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]


# Actuals Functions

def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(hexKeyCode, 0x48, 0, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(hexKeyCode, 0x48, 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def do_rotate(count):
    global example

    if example == "log":
        if count > 0:
            print("clockwise rotation of %d Detents", count)
        else:
            print("counterclockwise rotation of %d Detents", -count)
    elif example == "scroll":
        win32api.mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -count * 15, 0)
    elif example == "alttab":
        if count < 0:
            count = -count
            PressKey(0x10)  # shift
        for i in range(0, count):
            PressKey(0x09)  # Tab
            ReleaseKey(0x09)  # ~Tab
        ReleaseKey(0x10)  # ~shift


def do_click(down):
    global example
    if example == "log":
        if (down):
            print("button down")
        else:
            print("button up")
    elif example == "scroll":
        return
    elif example == "alttab":
        if down:
            PressKey(0x5B)  # Alt
            PressKey(0x09)  # Tab
        else:
            ReleaseKey(0x5B)  # ~Alt


def handleRotate(button, value):
    """
    handle rotation of encoder

    @type button: YAnButton
    @type value: str
    """
    global last_ev_pressed
    pressed = int(value) < 500
    if pressed != button.get_userData():
        if last_ev_pressed != pressed:
            if (button.get_logicalName() == "encoderA"):
                do_rotate(1)
            else:
                do_rotate(-1)
            last_ev_pressed = pressed
        button.set_userData(pressed)


def handleClick(button, value):
    """
    handle click button of encoder

    @type button: YAnButton
    @type value: str
    """
    pressed = button.get_isPressed()
    if pressed != button.get_userData():
        do_click(pressed)
        button.set_userData(pressed)


def handleRotatePolling():
    if encoderA.get_isPressed() != encoderB.get_isPressed():
        return
    counter = encoderA.get_pulseCounter()
    global old_counter
    nb_transitions = counter - old_counter
    if nb_transitions != 0:
        b_get_last_time_pressed = encoderB.get_lastTimePressed()
        a_get_last_time_pressed = encoderA.get_lastTimePressed()
        if (b_get_last_time_pressed > a_get_last_time_pressed):
            do_rotate(nb_transitions)
        else:
            do_rotate(-nb_transitions)
        old_counter = counter


errmsg = YRefParam()


# Setup the API to use local USB devices
if YAPI.RegisterHub("usb", errmsg) != YAPI.SUCCESS:
    sys.exit("init error" + errmsg.value)

print('Hit Ctrl-C to Stop ')

example = 'log'
#example = 'alttab'

encoderA = YAnButton.FindAnButton("encoderA")
encoderB = YAnButton.FindAnButton("encoderB")
encoderC = YAnButton.FindAnButton("encoderC")

if not encoderA.isOnline() or not encoderB.isOnline() or not encoderC.isOnline():
    sys.exit("no valid Yocto-Knob detected.")
encoderC.set_userData(encoderC.get_isPressed())

encoderA.set_pulseCounter(0)
encoderB.set_pulseCounter(0)

old_counter = encoderA.get_pulseCounter()

last_ev_pressed = encoderA.get_isPressed()
encoderA.registerValueCallback(handleRotate)
encoderB.registerValueCallback(handleRotate)
encoderC.registerValueCallback(handleClick)

while True:
    YAPI.Sleep(100, errmsg)  # traps others events

