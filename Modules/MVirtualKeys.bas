Attribute VB_Name = "MVirtualKeys"
Option Explicit

Public Enum EVirtualKeyCodes
    VK_LBUTTON = &H1         'Left mouse button
    VK_RBUTTON = &H2         'Right mouse button
    VK_CANCEL = &H3          'Control-break processing
    VK_MBUTTON = &H4         'Middle mouse button
    VK_XBUTTON1 = &H5        'X1 mouse button
    VK_XBUTTON2 = &H6        'X2 mouse button
                             '0x07 'Reserved
    VK_BACK = &H8            'Backspace key
    VK_TAB = &H9             'Tab key
            '0x0A-0B         'Reserved
    VK_CLEAR = &HC           'Clear key
    VK_RETURN = &HD          'Enter key
               '0x0E-0F      'Unassigned
    VK_SHIFT = &H10          'Shift key
    VK_CONTROL = &H11        'Ctrl key
    VK_MENU = &H12           'Alt key
    'VK_ALTGR = &H11 + &H12 '= 17 + 18
    VK_PAUSE = &H13          'Pause key
    VK_CAPITAL = &H14        'Caps lock key
    VK_KANA = &H15           'IME Kana mode
    VK_HANGUL = &H15         'IME Hangul mode
    VK_IME_ON = &H16         'IME On
    VK_JUNJA = &H17          'IME Junja mode
    VK_FINAL = &H18          'IME final mode
    VK_HANJA = &H19          'IME Hanja mode
    VK_KANJI = &H19          'IME Kanji mode
    VK_IME_OFF = &H1A        'IME Off
    VK_ESCAPE = &H1B         'Esc key
    VK_CONVERT = &H1C        'IME convert
    VK_NONCONVERT = &H1D     'IME nonconvert
    VK_ACCEPT = &H1E         'IME accept
    VK_MODECHANGE = &H1F     'IME mode change request
    VK_SPACE = &H20          'Spacebar key
    VK_PRIOR = &H21          'Page up key
    VK_NEXT = &H22           'Page down key
    VK_END = &H23            'End key
    VK_HOME = &H24           'Home key
    VK_LEFT = &H25           'Left arrow key
    VK_UP = &H26             'Up arrow key
    VK_RIGHT = &H27          'Right arrow key
    VK_DOWN = &H28           'Down arrow key
    VK_SELECT = &H29         'Select key
    VK_PRINT = &H2A          'Print key
    VK_EXECUTE = &H2B        'Execute key
    VK_SNAPSHOT = &H2C       'Print screen key
    VK_INSERT = &H2D         'Insert key
    VK_DELETE = &H2E         'Delete key
    VK_HELP = &H2F           'Help key
    VK_KEY_0 = &H30          '0 key
    VK_KEY_1 = &H31          '1 key
    VK_KEY_2 = &H32          '2 key
    VK_KEY_3 = &H33          '3 key
    VK_KEY_4 = &H34          '4 key
    VK_KEY_5 = &H35          '5 key
    VK_KEY_6 = &H36          '6 key
    VK_KEY_7 = &H37          '7 key
    VK_KEY_8 = &H38          '8 key
    VK_KEY_9 = &H39          '9 key
              '0x3A-40       'Undefined
    VK_KEY_A = &H41          'A key
    VK_KEY_B = &H42          'B key
    VK_KEY_C = &H43          'C key
    VK_KEY_D = &H44          'D key
    VK_KEY_E = &H45          'E key
    VK_KEY_F = &H46          'F key
    VK_KEY_G = &H47          'G key
    VK_KEY_H = &H48          'H key
    VK_KEY_I = &H49          'I key
    VK_KEY_J = &H4A          'J key
    VK_KEY_K = &H4B          'K key
    VK_KEY_L = &H4C          'L key
    VK_KEY_M = &H4D          'M key
    VK_KEY_N = &H4E          'N key
    VK_KEY_O = &H4F          'O key
    VK_KEY_P = &H50          'P key
    VK_KEY_Q = &H51          'Q key
    VK_KEY_R = &H52          'R key
    VK_KEY_S = &H53          'S key
    VK_KEY_T = &H54          'T key
    VK_KEY_U = &H55          'U key
    VK_KEY_V = &H56          'V key
    VK_KEY_W = &H57          'W key
    VK_KEY_X = &H58          'X key
    VK_KEY_Y = &H59          'Y key
    VK_KEY_Z = &H5A          'Z key
    VK_LWIN = &H5B           'Left Windows logo key
    VK_RWIN = &H5C           'Right Windows logo key
    VK_APPS = &H5D           'Application key
             '0x5E           'Reserved
    VK_SLEEP = &H5F          'Computer Sleep key
    VK_NUMPAD0 = &H60        'Numeric keypad 0 key
    VK_NUMPAD1 = &H61        'Numeric keypad 1 key
    VK_NUMPAD2 = &H62        'Numeric keypad 2 key
    VK_NUMPAD3 = &H63        'Numeric keypad 3 key
    VK_NUMPAD4 = &H64        'Numeric keypad 4 key
    VK_NUMPAD5 = &H65        'Numeric keypad 5 key
    VK_NUMPAD6 = &H66        'Numeric keypad 6 key
    VK_NUMPAD7 = &H67        'Numeric keypad 7 key
    VK_NUMPAD8 = &H68        'Numeric keypad 8 key
    VK_NUMPAD9 = &H69        'Numeric keypad 9 key
    VK_MULTIPLY = &H6A       'Multiply key
    VK_ADD = &H6B            'Add key
    VK_SEPARATOR = &H6C      'Separator key
    VK_SUBTRACT = &H6D       'Subtract key
    VK_DECIMAL = &H6E        'Decimal key
    VK_DIVIDE = &H6F         'Divide key
    VK_F1 = &H70             'F1 key
    VK_F2 = &H71             'F2 key
    VK_F3 = &H72             'F3 key
    VK_F4 = &H73             'F4 key
    VK_F5 = &H74             'F5 key
    VK_F6 = &H75             'F6 key
    VK_F7 = &H76             'F7 key
    VK_F8 = &H77             'F8 key
    VK_F9 = &H78             'F9 key
    VK_F10 = &H79            'F10 key
    VK_F11 = &H7A            'F11 key
    VK_F12 = &H7B            'F12 key
    VK_F13 = &H7C            'F13 key
    VK_F14 = &H7D            'F14 key
    VK_F15 = &H7E            'F15 key
    VK_F16 = &H7F            'F16 key
    VK_F17 = &H80            'F17 key
    VK_F18 = &H81            'F18 key
    VK_F19 = &H82            'F19 key
    VK_F20 = &H83            'F20 key
    VK_F21 = &H84            'F21 key
    VK_F22 = &H85            'F22 key
    VK_F23 = &H86            'F23 key
    VK_F24 = &H87            'F24 key
                '0x88-8F     'Reserved
    VK_NUMLOCK = &H90        'Num lock key
    VK_SCROLL = &H91         'Scroll lock key
               '0x92-96      'OEM specific
               '0x97-9F      'Unassigned
    VK_LSHIFT = &HA0         'Left Shift key
    VK_RSHIFT = &HA1         'Right Shift key
    VK_LCONTROL = &HA2       'Left Ctrl key
    VK_RCONTROL = &HA3       'Right Ctrl key
    VK_LMENU = &HA4          'Left Alt key
    VK_RMENU = &HA5          'Right Alt key
    VK_BROWSER_BACK = &HA6       'Browser Back key
    VK_BROWSER_FORWARD = &HA7    'Browser Forward key
    VK_BROWSER_REFRESH = &HA8    'Browser Refresh key
    VK_BROWSER_STOP = &HA9       'Browser Stop key
    VK_BROWSER_SEARCH = &HAA     'Browser Search key
    VK_BROWSER_FAVORITES = &HAB  'Browser Favorites key
    VK_BROWSER_HOME = &HAC       'Browser Start and Home key
    VK_VOLUME_MUTE = &HAD        'Volume Mute key
    VK_VOLUME_DOWN = &HAE        'Volume Down key
    VK_VOLUME_UP = &HAF          'Volume Up key
    VK_MEDIA_NEXT_TRACK = &HB0   'Next Track key
    VK_MEDIA_PREV_TRACK = &HB1   'Previous Track key
    VK_MEDIA_STOP = &HB2         'Stop Media key
    VK_MEDIA_PLAY_PAUSE = &HB3   'Play/Pause Media key
    VK_LAUNCH_MAIL = &HB4        'Start Mail key
    VK_LAUNCH_MEDIA_SELECT = &HB5    'Select Media key
    VK_LAUNCH_APP1 = &HB6        'Start Application 1 key
    VK_LAUNCH_APP2 = &HB7        'Start Application 2 key
                    '0xB8-B9     'Reserved
    VK_OEM_1 = &HBA              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ;: key
    VK_OEM_PLUS = &HBB           'For any country/region, the + key
    VK_OEM_COMMA = &HBC          'For any country/region, the , key
    VK_OEM_MINUS = &HBD          'For any country/region, the - key
    VK_OEM_PERIOD = &HBE         'For any country/region, the . key
    VK_OEM_2 = &HBF              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the /? key
    VK_OEM_3 = &HC0              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the `~ key
              '0xC1-DA           'Reserved
    VK_OEM_4 = &HDB              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the [{ key
    VK_OEM_5 = &HDC              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the \\| key
    VK_OEM_6 = &HDD              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ]} key
    VK_OEM_7 = &HDE              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '" key
    VK_OEM_8 = &HDF              'Used for miscellaneous characters; it can vary by keyboard.
              '0xE0              'Reserved
              '0xE1              'OEM specific
    VK_OEM_102 = &HE2            'The <> keys on the US standard keyboard, or the \\| key on the non-US 102-key keyboard
                '0xE3-E4         'OEM specific
    VK_PROCESSKEY = &HE5         'IME PROCESS key
                   '0xE6         'OEM specific
    VK_PACKET = &HE7             'Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
               '0xE8             'Unassigned
        '0xE9-F5                 'OEM specific
    VK_ATTN = &HF6               'Attn key
    VK_CRSEL = &HF7              'CrSel key
    VK_EXSEL = &HF8              'ExSel key
    VK_EREOF = &HF9              'Erase EOF key
    VK_PLAY = &HFA               'Play key
    VK_ZOOM = &HFB               'Zoom key
    VK_NONAME = &HFC             'Reserved
    VK_PA1 = &HFD                'PA1 key
    VK_OEM_CLEAR = &HFE          'Clear key
End Enum

Public Enum EKeyEventFlags
    KEYEVENTF_KEYDOWN = &H0&
    KEYEVENTF_EXTENDEDKEY = &H1&     '| If specified, the wScan scan code consists of a sequence of two bytes, where the first byte has a value of 0xE0. See Extended-Key Flag for more info.
    KEYEVENTF_KEYUP = &H2&           '| If specified, the key is being released. If not specified, the key is being pressed.
    KEYEVENTF_UNICODE = &H4&         '| If specified, the system synthesizes a VK_PACKET keystroke. The wVk parameter must be zero. This flag can only be combined with the KEYEVENTF_KEYUP flag. For more information, see the Remarks section.
    KEYEVENTF_SCANCODE = &H8&        '| If specified, wScan identifies the key and wVk is ignored.
End Enum

Public Enum EMouseEventFlags
    MOUSEEVENTF_MOVE = &H1&                ' 2^ 0 'Movement occurred.
    MOUSEEVENTF_LEFTDOWN = &H2&            ' 2^ 1 'The left button was pressed.
    MOUSEEVENTF_LEFTUP = &H4&              ' 2^ 2 'The left button was released.
    MOUSEEVENTF_RIGHTDOWN = &H8&           ' 2^ 3 'The right button was pressed.
    MOUSEEVENTF_RIGHTUP = &H10&            ' 2^ 4 'The right button was released.
    MOUSEEVENTF_MIDDLEDOWN = &H20&         ' 2^ 5 'The middle button was pressed.
    MOUSEEVENTF_MIDDLEUP = &H40&           ' 2^ 6 'The middle button was released.
    MOUSEEVENTF_XDOWN = &H80&              ' 2^ 7 'An X button was pressed.
    MOUSEEVENTF_XUP = &H100&               ' 2^ 8 'An X button was released.
    ' 200                                  ' 2^ 9
    ' 400                                  ' 2^10
    MOUSEEVENTF_WHEEL = &H800&             ' 2^11 'The wheel was moved, if the mouse has a wheel. The amount of movement is specified in mouseData.
    MOUSEEVENTF_HWHEEL = &H1000&           ' 2^12 'The wheel was moved horizontally, if the mouse has a wheel. The amount of movement is specified in mouseData. Windows XP/2000: This value is not supported.
    MOUSEEVENTF_MOVE_NOCOALESCE = &H2000&  ' 2^13 'The WM_MOUSEMOVE messages will not be coalesced. The default behavior is to coalesce WM_MOUSEMOVE messages.Windows XP/2000: This value is not supported.
    MOUSEEVENTF_VIRTUALDESK = &H4000&      ' 2^14 'Maps coordinates to the entire desktop. Must be used with MOUSEEVENTF_ABSOLUTE.
    MOUSEEVENTF_ABSOLUTE = &H8000&         ' 2^15 'The dx and dy members contain normalized absolute coordinates. If the flag is not set, dxand dy contain relative data (the change in position since the last reported position). This flag can be set, or not set, regardless of what kind of mouse or other pointing device, if any, is connected to the system. For further information about relative mouse motion, see the following Remarks section.
End Enum
Private m_VKeyNmInit  As Boolean
Private m_VKeyNames() As String

' v ' ############################## ' v '    EVirtualKeyCodes     ' v ' ############################## ' v '
Public Function EVirtualKeyCodes_ToStr(ByVal e As EVirtualKeyCodes) As String
    Dim s As String
    Select Case e
    Case 0: s = ""
    Case EVirtualKeyCodes.VK_LBUTTON:             s = "VK_LBUTTON" ' &H1         'Left mouse button
    Case EVirtualKeyCodes.VK_RBUTTON:             s = "VK_RBUTTON" ' &H2             'Right mouse button
    Case EVirtualKeyCodes.VK_CANCEL:              s = "VK_CANCEL" ' &H3          'Control-break processing
    Case EVirtualKeyCodes.VK_MBUTTON:             s = "VK_MBUTTON" ' &H4         'Middle mouse button
    Case EVirtualKeyCodes.VK_XBUTTON1:            s = "VK_XBUTTON1" ' &H5        'X1 mouse button
    Case EVirtualKeyCodes.VK_XBUTTON2:            s = "VK_XBUTTON2" ' &H6        'X2 mouse button
                             '0x07 'Reserved
    Case EVirtualKeyCodes.VK_BACK:                s = "VK_BACK" ' &H8            'Backspace key
    Case EVirtualKeyCodes.VK_TAB:                 s = "VK_TAB" ' &H9             'Tab key
            '0x0A-0B         'Reserved
    Case EVirtualKeyCodes.VK_CLEAR:               s = "VK_CLEAR" ' &HC           'Clear key
    Case EVirtualKeyCodes.VK_RETURN:              s = "VK_RETURN" ' &HD          'Enter key
               '0x0E-0F      'Unassigned
    Case EVirtualKeyCodes.VK_SHIFT:               s = "VK_SHIFT" ' &H10          'Shift key
    Case EVirtualKeyCodes.VK_CONTROL:             s = "VK_CONTROL" ' &H11        'Ctrl key
    Case EVirtualKeyCodes.VK_MENU:                s = "VK_MENU" ' &H12           'Alt key
    Case EVirtualKeyCodes.VK_PAUSE:               s = "VK_PAUSE" ' &H13          'Pause key
    Case EVirtualKeyCodes.VK_CAPITAL:             s = "VK_CAPITAL" ' &H14        'Caps lock key
    Case EVirtualKeyCodes.VK_KANA:                s = "VK_KANA" ' &H15           'IME Kana mode
    Case EVirtualKeyCodes.VK_HANGUL:              s = "VK_HANGUL" ' &H15         'IME Hangul mode
    Case EVirtualKeyCodes.VK_IME_ON:              s = "VK_IME_ON" ' &H16         'IME On
    Case EVirtualKeyCodes.VK_JUNJA:               s = "VK_JUNJA" ' &H17          'IME Junja mode
    Case EVirtualKeyCodes.VK_FINAL:               s = "VK_FINAL" ' &H18          'IME final mode
    Case EVirtualKeyCodes.VK_HANJA:               s = "VK_HANJA" ' &H19          'IME Hanja mode
    Case EVirtualKeyCodes.VK_KANJI:               s = "VK_KANJI" ' &H19          'IME Kanji mode
    Case EVirtualKeyCodes.VK_IME_OFF:             s = "VK_IME_OFF" ' &H1A        'IME Off
    Case EVirtualKeyCodes.VK_ESCAPE:              s = "VK_ESCAPE" ' &H1B         'Esc key
    Case EVirtualKeyCodes.VK_CONVERT:             s = "VK_CONVERT" ' &H1C        'IME convert
    Case EVirtualKeyCodes.VK_NONCONVERT:          s = "VK_NONCONVERT" ' &H1D     'IME nonconvert
    Case EVirtualKeyCodes.VK_ACCEPT:              s = "VK_ACCEPT" ' &H1E         'IME accept
    Case EVirtualKeyCodes.VK_MODECHANGE:          s = "VK_MODECHANGE" ' &H1F     'IME mode change request
    Case EVirtualKeyCodes.VK_SPACE:               s = "VK_SPACE" ' &H20          'Spacebar key
    Case EVirtualKeyCodes.VK_PRIOR:               s = "VK_PRIOR" ' &H21          'Page up key
    Case EVirtualKeyCodes.VK_NEXT:                s = "VK_NEXT" ' &H22           'Page down key
    Case EVirtualKeyCodes.VK_END:                 s = "VK_END" ' &H23            'End key
    Case EVirtualKeyCodes.VK_HOME:                s = "VK_HOME" ' &H24           'Home key
    Case EVirtualKeyCodes.VK_LEFT:                s = "VK_LEFT" ' &H25           'Left arrow key
    Case EVirtualKeyCodes.VK_UP:                  s = "VK_UP" ' &H26             'Up arrow key
    Case EVirtualKeyCodes.VK_RIGHT:               s = "VK_RIGHT" ' &H27          'Right arrow key
    Case EVirtualKeyCodes.VK_DOWN:                s = "VK_DOWN" ' &H28           'Down arrow key
    Case EVirtualKeyCodes.VK_SELECT:              s = "VK_SELECT" ' &H29         'Select key
    Case EVirtualKeyCodes.VK_PRINT:               s = "VK_PRINT" ' &H2A          'Print key
    Case EVirtualKeyCodes.VK_EXECUTE:             s = "VK_EXECUTE" ' &H2B        'Execute key
    Case EVirtualKeyCodes.VK_SNAPSHOT:            s = "VK_SNAPSHOT" ' &H2C       'Print screen key
    Case EVirtualKeyCodes.VK_INSERT:              s = "VK_INSERT" ' &H2D         'Insert key
    Case EVirtualKeyCodes.VK_DELETE:              s = "VK_DELETE" ' &H2E         'Delete key
    Case EVirtualKeyCodes.VK_HELP:                s = "VK_HELP" ' &H2F           'Help key
    Case EVirtualKeyCodes.VK_KEY_0:               s = "VK_KEY_0" ' &H30          '0 key
    Case EVirtualKeyCodes.VK_KEY_1:               s = "VK_KEY_1" ' &H31          '1 key
    Case EVirtualKeyCodes.VK_KEY_2:               s = "VK_KEY_2" ' &H32          '2 key
    Case EVirtualKeyCodes.VK_KEY_3:               s = "VK_KEY_3" ' &H33          '3 key
    Case EVirtualKeyCodes.VK_KEY_4:               s = "VK_KEY_4" ' &H34          '4 key
    Case EVirtualKeyCodes.VK_KEY_5:               s = "VK_KEY_5" ' &H35          '5 key
    Case EVirtualKeyCodes.VK_KEY_6:               s = "VK_KEY_6" ' &H36          '6 key
    Case EVirtualKeyCodes.VK_KEY_7:               s = "VK_KEY_7" ' &H37          '7 key
    Case EVirtualKeyCodes.VK_KEY_8:               s = "VK_KEY_8" ' &H38          '8 key
    Case EVirtualKeyCodes.VK_KEY_9:               s = "VK_KEY_9" ' &H39          '9 key
               '0x3A-40       'Undefined
    Case EVirtualKeyCodes.VK_KEY_A:               s = "VK_KEY_A" ' &H41          'A key
    Case EVirtualKeyCodes.VK_KEY_B:               s = "VK_KEY_B" ' &H42          'B key
    Case EVirtualKeyCodes.VK_KEY_C:               s = "VK_KEY_C" ' &H43          'C key
    Case EVirtualKeyCodes.VK_KEY_D:               s = "VK_KEY_D" ' &H44          'D key
    Case EVirtualKeyCodes.VK_KEY_E:               s = "VK_KEY_E" ' &H45          'E key
    Case EVirtualKeyCodes.VK_KEY_F:               s = "VK_KEY_F" ' &H46          'F key
    Case EVirtualKeyCodes.VK_KEY_G:               s = "VK_KEY_G" ' &H47          'G key
    Case EVirtualKeyCodes.VK_KEY_H:               s = "VK_KEY_H" ' &H48          'H key
    Case EVirtualKeyCodes.VK_KEY_I:               s = "VK_KEY_I" ' &H49          'I key
    Case EVirtualKeyCodes.VK_KEY_J:               s = "VK_KEY_J" ' &H4A          'J key
    Case EVirtualKeyCodes.VK_KEY_K:               s = "VK_KEY_K" ' &H4B          'K key
    Case EVirtualKeyCodes.VK_KEY_L:               s = "VK_KEY_L" ' &H4C          'L key
    Case EVirtualKeyCodes.VK_KEY_M:               s = "VK_KEY_M" ' &H4D          'M key
    Case EVirtualKeyCodes.VK_KEY_N:               s = "VK_KEY_N" ' &H4E          'N key
    Case EVirtualKeyCodes.VK_KEY_O:               s = "VK_KEY_O" ' &H4F          'O key
    Case EVirtualKeyCodes.VK_KEY_P:               s = "VK_KEY_P" ' &H50          'P key
    Case EVirtualKeyCodes.VK_KEY_Q:               s = "VK_KEY_Q" ' &H51          'Q key
    Case EVirtualKeyCodes.VK_KEY_R:               s = "VK_KEY_R" ' &H52          'R key
    Case EVirtualKeyCodes.VK_KEY_S:               s = "VK_KEY_S" ' &H53          'S key
    Case EVirtualKeyCodes.VK_KEY_T:               s = "VK_KEY_T" ' &H54          'T key
    Case EVirtualKeyCodes.VK_KEY_U:               s = "VK_KEY_U" ' &H55          'U key
    Case EVirtualKeyCodes.VK_KEY_V:               s = "VK_KEY_V" ' &H56          'V key
    Case EVirtualKeyCodes.VK_KEY_W:               s = "VK_KEY_W" ' &H57          'W key
    Case EVirtualKeyCodes.VK_KEY_X:               s = "VK_KEY_X" ' &H58          'X key
    Case EVirtualKeyCodes.VK_KEY_Y:               s = "VK_KEY_Y" ' &H59          'Y key
    Case EVirtualKeyCodes.VK_KEY_Z:               s = "VK_KEY_Z" ' &H5A          'Z key
    Case EVirtualKeyCodes.VK_LWIN:                s = "VK_LWIN" ' &H5B           'Left Windows logo key
    Case EVirtualKeyCodes.VK_RWIN:                s = "VK_RWIN" ' &H5C           'Right Windows logo key
    Case EVirtualKeyCodes.VK_APPS:                s = "VK_APPS" ' &H5D           'Application key
             '0x5E           'Reserved
    Case EVirtualKeyCodes.VK_SLEEP:               s = "VK_SLEEP" ' &H5F          'Computer Sleep key
    Case EVirtualKeyCodes.VK_NUMPAD0:             s = "VK_NUMPAD0" ' &H60        'Numeric keypad 0 key
    Case EVirtualKeyCodes.VK_NUMPAD1:             s = "VK_NUMPAD1" ' &H61        'Numeric keypad 1 key
    Case EVirtualKeyCodes.VK_NUMPAD2:             s = "VK_NUMPAD2" ' &H62        'Numeric keypad 2 key
    Case EVirtualKeyCodes.VK_NUMPAD3:             s = "VK_NUMPAD3" ' &H63        'Numeric keypad 3 key
    Case EVirtualKeyCodes.VK_NUMPAD4:             s = "VK_NUMPAD4" ' &H64        'Numeric keypad 4 key
    Case EVirtualKeyCodes.VK_NUMPAD5:             s = "VK_NUMPAD5" ' &H65        'Numeric keypad 5 key
    Case EVirtualKeyCodes.VK_NUMPAD6:             s = "VK_NUMPAD6" ' &H66        'Numeric keypad 6 key
    Case EVirtualKeyCodes.VK_NUMPAD7:             s = "VK_NUMPAD7" ' &H67        'Numeric keypad 7 key
    Case EVirtualKeyCodes.VK_NUMPAD8:             s = "VK_NUMPAD8" ' &H68        'Numeric keypad 8 key
    Case EVirtualKeyCodes.VK_NUMPAD9:             s = "VK_NUMPAD9" ' &H69        'Numeric keypad 9 key
    Case EVirtualKeyCodes.VK_MULTIPLY:            s = "VK_MULTIPLY" ' &H6A       'Multiply key
    Case EVirtualKeyCodes.VK_ADD:                 s = "VK_ADD" ' &H6B            'Add key
    Case EVirtualKeyCodes.VK_SEPARATOR:           s = "VK_SEPARATOR" ' &H6C      'Separator key
    Case EVirtualKeyCodes.VK_SUBTRACT:            s = "VK_SUBTRACT" ' &H6D       'Subtract key
    Case EVirtualKeyCodes.VK_DECIMAL:             s = "VK_DECIMAL" ' &H6E        'Decimal key
    Case EVirtualKeyCodes.VK_DIVIDE:              s = "VK_DIVIDE" ' &H6F         'Divide key
    Case EVirtualKeyCodes.VK_F1:                  s = "VK_F1" ' &H70             'F1 key
    Case EVirtualKeyCodes.VK_F2:                  s = "VK_F2" ' &H71             'F2 key
    Case EVirtualKeyCodes.VK_F3:                  s = "VK_F3" ' &H72             'F3 key
    Case EVirtualKeyCodes.VK_F4:                  s = "VK_F4" ' &H73             'F4 key
    Case EVirtualKeyCodes.VK_F5:                  s = "VK_F5" ' &H74             'F5 key
    Case EVirtualKeyCodes.VK_F6:                  s = "VK_F6" ' &H75             'F6 key
    Case EVirtualKeyCodes.VK_F7:                  s = "VK_F7" ' &H76             'F7 key
    Case EVirtualKeyCodes.VK_F8:                  s = "VK_F8" ' &H77             'F8 key
    Case EVirtualKeyCodes.VK_F9:                  s = "VK_F9" ' &H78             'F9 key
    Case EVirtualKeyCodes.VK_F10:                 s = "VK_F10" ' &H79            'F10 key
    Case EVirtualKeyCodes.VK_F11:                 s = "VK_F11" ' &H7A            'F11 key
    Case EVirtualKeyCodes.VK_F12:                 s = "VK_F12" ' &H7B            'F12 key
    Case EVirtualKeyCodes.VK_F13:                 s = "VK_F13" ' &H7C            'F13 key
    Case EVirtualKeyCodes.VK_F14:                 s = "VK_F14" ' &H7D            'F14 key
    Case EVirtualKeyCodes.VK_F15:                 s = "VK_F15" ' &H7E            'F15 key
    Case EVirtualKeyCodes.VK_F16:                 s = "VK_F16" ' &H7F            'F16 key
    Case EVirtualKeyCodes.VK_F17:                 s = "VK_F17" ' &H80            'F17 key
    Case EVirtualKeyCodes.VK_F18:                 s = "VK_F18" ' &H81            'F18 key
    Case EVirtualKeyCodes.VK_F19:                 s = "VK_F19" ' &H82            'F19 key
    Case EVirtualKeyCodes.VK_F20:                 s = "VK_F20" ' &H83            'F20 key
    Case EVirtualKeyCodes.VK_F21:                 s = "VK_F21" ' &H84            'F21 key
    Case EVirtualKeyCodes.VK_F22:                 s = "VK_F22" ' &H85            'F22 key
    Case EVirtualKeyCodes.VK_F23:                 s = "VK_F23" ' &H86            'F23 key
    Case EVirtualKeyCodes.VK_F24:                 s = "VK_F24" ' &H87            'F24 key
                '0x88-8F     'Reserved
    Case EVirtualKeyCodes.VK_NUMLOCK:             s = "VK_NUMLOCK" ' &H90        'Num lock key
    Case EVirtualKeyCodes.VK_SCROLL:              s = "VK_SCROLL" ' &H91         'Scroll lock key
               '0x92-96      'OEM specific
               '0x97-9F      'Unassigned
    Case EVirtualKeyCodes.VK_LSHIFT:              s = "VK_LSHIFT" ' &HA0         'Left Shift key
    Case EVirtualKeyCodes.VK_RSHIFT:              s = "VK_RSHIFT" ' &HA1         'Right Shift key
    Case EVirtualKeyCodes.VK_LCONTROL:            s = "VK_LCONTROL" ' &HA2       'Left Ctrl key
    Case EVirtualKeyCodes.VK_RCONTROL:            s = "VK_RCONTROL" ' &HA3       'Right Ctrl key
    Case EVirtualKeyCodes.VK_LMENU:               s = "VK_LMENU" ' &HA4          'Left Alt key
    Case EVirtualKeyCodes.VK_RMENU:               s = "VK_RMENU" ' &HA5          'Right Alt key
    Case EVirtualKeyCodes.VK_BROWSER_BACK:        s = "VK_BROWSER_BACK" ' &HA6       'Browser Back key
    Case EVirtualKeyCodes.VK_BROWSER_FORWARD:     s = "VK_BROWSER_FORWARD" ' &HA7    'Browser Forward key
    Case EVirtualKeyCodes.VK_BROWSER_REFRESH:     s = "VK_BROWSER_REFRESH" ' &HA8    'Browser Refresh key
    Case EVirtualKeyCodes.VK_BROWSER_STOP:        s = "VK_BROWSER_STOP" ' &HA9       'Browser Stop key
    Case EVirtualKeyCodes.VK_BROWSER_SEARCH:      s = "VK_BROWSER_SEARCH" ' &HAA     'Browser Search key
    Case EVirtualKeyCodes.VK_BROWSER_FAVORITES:   s = "VK_BROWSER_FAVORITES" ' &HAB  'Browser Favorites key
    Case EVirtualKeyCodes.VK_BROWSER_HOME:        s = "VK_BROWSER_HOME" ' &HAC       'Browser Start and Home key
    Case EVirtualKeyCodes.VK_VOLUME_MUTE:         s = "VK_VOLUME_MUTE" ' &HAD        'Volume Mute key
    Case EVirtualKeyCodes.VK_VOLUME_DOWN:         s = "VK_VOLUME_DOWN" ' &HAE        'Volume Down key
    Case EVirtualKeyCodes.VK_VOLUME_UP:           s = "VK_VOLUME_UP" ' &HAF          'Volume Up key
    Case EVirtualKeyCodes.VK_MEDIA_NEXT_TRACK:    s = "VK_MEDIA_NEXT_TRACK" ' &HB0   'Next Track key
    Case EVirtualKeyCodes.VK_MEDIA_PREV_TRACK:    s = "VK_MEDIA_PREV_TRACK" ' &HB1   'Previous Track key
    Case EVirtualKeyCodes.VK_MEDIA_STOP:          s = "VK_MEDIA_STOP" ' &HB2         'Stop Media key
    Case EVirtualKeyCodes.VK_MEDIA_PLAY_PAUSE:    s = "VK_MEDIA_PLAY_PAUSE" ' &HB3   'Play/Pause Media key
    Case EVirtualKeyCodes.VK_LAUNCH_MAIL:         s = "VK_LAUNCH_MAIL" ' &HB4        'Start Mail key
    Case EVirtualKeyCodes.VK_LAUNCH_MEDIA_SELECT: s = "VK_LAUNCH_MEDIA_SELECT" ' &HB5    'Select Media key
    Case EVirtualKeyCodes.VK_LAUNCH_APP1:         s = "VK_LAUNCH_APP1" ' &HB6        'Start Application 1 key
    Case EVirtualKeyCodes.VK_LAUNCH_APP2:         s = "VK_LAUNCH_APP2" ' &HB7        'Start Application 2 key
                    '0xB8-B9     'Reserved
    Case EVirtualKeyCodes.VK_OEM_1:               s = "VK_OEM_1" ' &HBA              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ;: key
    Case EVirtualKeyCodes.VK_OEM_PLUS:            s = "VK_OEM_PLUS" ' &HBB           'For any country/region, the + key
    Case EVirtualKeyCodes.VK_OEM_COMMA:           s = "VK_OEM_COMMA" ' &HBC          'For any country/region, the , key
    Case EVirtualKeyCodes.VK_OEM_MINUS:           s = "VK_OEM_MINUS" ' &HBD          'For any country/region, the - key
    Case EVirtualKeyCodes.VK_OEM_PERIOD:          s = "VK_OEM_PERIOD" ' &HBE         'For any country/region, the . key
    Case EVirtualKeyCodes.VK_OEM_2:               s = "VK_OEM_2" ' &HBF              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the /? key
    Case EVirtualKeyCodes.VK_OEM_3:               s = "VK_OEM_3" ' &HC0              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the `~ key
              '0xC1-DA           'Reserved
    Case EVirtualKeyCodes.VK_OEM_4:               s = "VK_OEM_4" ' &HDB              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the [{ key
    Case EVirtualKeyCodes.VK_OEM_5:               s = "VK_OEM_5" ' &HDC              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the \\| key
    Case EVirtualKeyCodes.VK_OEM_6:               s = "VK_OEM_6" ' &HDD              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ]} key
    Case EVirtualKeyCodes.VK_OEM_7:               s = "VK_OEM_7" ' &HDE              'Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '" key
    Case EVirtualKeyCodes.VK_OEM_8:               s = "VK_OEM_8" ' &HDF              'Used for miscellaneous characters; it can vary by keyboard.
              '0xE0              'Reserved
              '0xE1              'OEM specific
    Case EVirtualKeyCodes.VK_OEM_102:             s = "VK_OEM_102" ' &HE2            'The <> keys on the US standard keyboard, or the \\| key on the non-US 102-key keyboard
                '0xE3-E4         'OEM specific
    Case EVirtualKeyCodes.VK_PROCESSKEY:          s = "VK_PROCESSKEY" ' &HE5         'IME PROCESS key
                   '0xE6         'OEM specific
    Case EVirtualKeyCodes.VK_PACKET:              s = "VK_PACKET" ' &HE7             'Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
               '0xE8             'Unassigned
        '0xE9-F5                 'OEM specific
    Case EVirtualKeyCodes.VK_ATTN:                s = "VK_ATTN" ' &HF6               'Attn key
    Case EVirtualKeyCodes.VK_CRSEL:               s = "VK_CRSEL" ' &HF7              'CrSel key
    Case EVirtualKeyCodes.VK_EXSEL:               s = "VK_EXSEL" ' &HF8              'ExSel key
    Case EVirtualKeyCodes.VK_EREOF:               s = "VK_EREOF" ' &HF9              'Erase EOF key
    Case EVirtualKeyCodes.VK_PLAY:                s = "VK_PLAY" ' &HFA               'Play key
    Case EVirtualKeyCodes.VK_ZOOM:                s = "VK_ZOOM" ' &HFB               'Zoom key
    Case EVirtualKeyCodes.VK_NONAME:              s = "VK_NONAME" ' &HFC             'Reserved
    Case EVirtualKeyCodes.VK_PA1:                 s = "VK_PA1" ' &HFD                'PA1 key
    Case EVirtualKeyCodes.VK_OEM_CLEAR:           s = "VK_OEM_CLEAR" ' &HFE          'Clear key
    Case Else: s = ""
    End Select
    EVirtualKeyCodes_ToStr = s
End Function

Public Sub EVirtualKeyCodes_ToList(aCmb As ComboBox)
    aCmb.Clear
    Dim i As Long, s As String
    aCmb.AddItem s
    If Not m_VKeyNmInit Then ReDim m_VKeyNames(0 To 300)
    For i = 0 To 300
        If m_VKeyNmInit Then
            s = m_VKeyNames(i)
        Else
            s = EVirtualKeyCodes_ToStr(i)
            m_VKeyNames(i) = s
        End If
        If Len(s) Then aCmb.AddItem s ' i & ", 0x" & Hex(i) & ", " & s
    Next
    m_VKeyNmInit = True
End Sub

Public Function EVirtualKeyCodes_Parse(ByVal s As String) As EVirtualKeyCodes
    Dim e As EVirtualKeyCodes ': e = -1 ' wieso -1?
    EVirtualKeyCodes_Parse = e
    If Left(s, 3) <> "VK_" Then Exit Function
    s = UCase(Mid(s, 4))
    Select Case s
    Case "LBUTTON":     e = &H1
    Case "RBUTTON":     e = &H2
    Case "CANCEL":      e = &H3
    Case "MBUTTON":     e = &H4
    Case "XBUTTON1":    e = &H5
    Case "XBUTTON2":    e = &H6
'#WERT!
    Case "BACK":        e = &H8
    Case "TAB":         e = &H9
'#WERT!
    Case "CLEAR":       e = &HC
    Case "RETURN":      e = &HD
'#WERT!
    Case "SHIFT":       e = &H10
    Case "CONTROL":     e = &H11
    Case "MENU":        e = &H12
    Case "PAUSE":       e = &H13
    Case "CAPITAL":     e = &H14
    Case "KANA":        e = &H15
    Case "HANGUL":              e = &H15
    Case "IME_ON":              e = &H16
    Case "JUNJA":               e = &H17
    Case "FINAL":               e = &H18
    Case "HANJA":               e = &H19
    Case "KANJI":               e = &H19
    Case "IME_OFF":             e = &H1A
    Case "ESCAPE":              e = &H1B
    Case "CONVERT":             e = &H1C
    Case "NONCONVERT":          e = &H1D
    Case "ACCEPT":              e = &H1E
    Case "MODECHANGE":          e = &H1F
    Case "SPACE":               e = &H20
    Case "PRIOR":               e = &H21
    Case "NEXT":                e = &H22
    Case "END":                 e = &H23
    Case "HOME":                e = &H24
    Case "LEFT":                e = &H25
    Case "UP":                  e = &H26
    Case "RIGHT":               e = &H27
    Case "DOWN":                e = &H28
    Case "SELECT":              e = &H29
    Case "PRINT":               e = &H2A
    Case "EXECUTE":             e = &H2B
    Case "SNAPSHOT":            e = &H2C
    Case "INSERT":              e = &H2D
    Case "DELETE":              e = &H2E
    Case "HELP":                e = &H2F
    Case "KEY_0":               e = &H30
    Case "KEY_1":               e = &H31
    Case "KEY_2":               e = &H32
    Case "KEY_3":               e = &H33
    Case "KEY_4":               e = &H34
    Case "KEY_5":               e = &H35
    Case "KEY_6":               e = &H36
    Case "KEY_7":               e = &H37
    Case "KEY_8":               e = &H38
    Case "KEY_9":               e = &H39
'#WERT!
    Case "KEY_A":               e = &H41
    Case "KEY_B":               e = &H42
    Case "KEY_C":               e = &H43
    Case "KEY_D":               e = &H44
    Case "KEY_E":               e = &H45
    Case "KEY_F":               e = &H46
    Case "KEY_G":               e = &H47
    Case "KEY_H":               e = &H48
    Case "KEY_I":               e = &H49
    Case "KEY_J":               e = &H4A
    Case "KEY_K":               e = &H4B
    Case "KEY_L":               e = &H4C
    Case "KEY_M":               e = &H4D
    Case "KEY_N":               e = &H4E
    Case "KEY_O":               e = &H4F
    Case "KEY_P":               e = &H50
    Case "KEY_Q":               e = &H51
    Case "KEY_R":               e = &H52
    Case "KEY_S":               e = &H53
    Case "KEY_T":               e = &H54
    Case "KEY_U":               e = &H55
    Case "KEY_V":               e = &H56
    Case "KEY_W":               e = &H57
    Case "KEY_X":               e = &H58
    Case "KEY_Y":               e = &H59
    Case "KEY_Z":               e = &H5A
    Case "LWIN":                e = &H5B
    Case "RWIN":                e = &H5C
    Case "APPS":                e = &H5D
'#WERT!
    Case "SLEEP":               e = &H5F
    Case "NUMPAD0":             e = &H60
    Case "NUMPAD1":             e = &H61
    Case "NUMPAD2":             e = &H62
    Case "NUMPAD3":             e = &H63
    Case "NUMPAD4":             e = &H64
    Case "NUMPAD5":             e = &H65
    Case "NUMPAD6":             e = &H66
    Case "NUMPAD7":             e = &H67
    Case "NUMPAD8":             e = &H68
    Case "NUMPAD9":             e = &H69
    Case "MULTIPLY":            e = &H6A
    Case "ADD":                 e = &H6B
    Case "SEPARATOR":           e = &H6C
    Case "SUBTRACT":            e = &H6D
    Case "DECIMAL":             e = &H6E
    Case "DIVIDE":              e = &H6F
    Case "F1":                  e = &H70
    Case "F2":                  e = &H71
    Case "F3":                  e = &H72
    Case "F4":                  e = &H73
    Case "F5":                  e = &H74
    Case "F6":                  e = &H75
    Case "F7":                  e = &H76
    Case "F8":                  e = &H77
    Case "F9":                  e = &H78
    Case "F10":                 e = &H79
    Case "F11":                 e = &H7A
    Case "F12":                 e = &H7B
    Case "F13":                 e = &H7C
    Case "F14":                 e = &H7D
    Case "F15":                 e = &H7E
    Case "F16":                 e = &H7F
    Case "F17":                 e = &H80
    Case "F18":                 e = &H81
    Case "F19":                 e = &H82
    Case "F20":                 e = &H83
    Case "F21":                 e = &H84
    Case "F22":                 e = &H85
    Case "F23":                 e = &H86
    Case "F24":                 e = &H87
'#WERT!
    Case "NUMLOCK":             e = &H90
    Case "SCROLL":              e = &H91
'#WERT!
'#WERT!
    Case "LSHIFT":              e = &HA0
    Case "RSHIFT":              e = &HA1
    Case "LCONTROL":            e = &HA2
    Case "RCONTROL":            e = &HA3
    Case "LMENU":               e = &HA4
    Case "RMENU":               e = &HA5
    Case "BROWSER_BACK":        e = &HA6
    Case "BROWSER_FORWARD":     e = &HA7
    Case "BROWSER_REFRESH":     e = &HA8
    Case "BROWSER_STOP":        e = &HA9
    Case "BROWSER_SEARCH":      e = &HAA
    Case "BROWSER_FAVORITES":   e = &HAB
    Case "BROWSER_HOME":        e = &HAC
    Case "VOLUME_MUTE":         e = &HAD
    Case "VOLUME_DOWN":         e = &HAE
    Case "VOLUME_UP":           e = &HAF
    Case "MEDIA_NEXT_TRACK":    e = &HB0
    Case "MEDIA_PREV_TRACK":    e = &HB1
    Case "MEDIA_STOP":          e = &HB2
    Case "MEDIA_PLAY_PAUSE":    e = &HB3
    Case "LAUNCH_MAIL":         e = &HB4
    Case "LAUNCH_MEDIA_SELECT": e = &HB5
    Case "LAUNCH_APP1":         e = &HB6
    Case "LAUNCH_APP2":         e = &HB7
'#WERT!
    Case "OEM_1":               e = &HBA
    Case "OEM_PLUS":            e = &HBB
    Case "OEM_COMMA":           e = &HBC
    Case "OEM_MINUS":           e = &HBD
    Case "OEM_PERIOD":          e = &HBE
    Case "OEM_2":               e = &HBF
    Case "OEM_3":               e = &HC0
'#WERT!
    Case "OEM_4":               e = &HDB
    Case "OEM_5":               e = &HDC
    Case "OEM_6":               e = &HDD
    Case "OEM_7":               e = &HDE
    Case "OEM_8":               e = &HDF
'#WERT!
'#WERT!
    Case "OEM_102":             e = &HE2
'#WERT!
    Case "PROCESSKEY":          e = &HE5
'#WERT!
    Case "PACKET":              e = &HE7
'#WERT!
'#WERT!
    Case "ATTN":                e = &HF6
    Case "CRSEL":               e = &HF7
    Case "EXSEL":               e = &HF8
    Case "EREOF":               e = &HF9
    Case "PLAY":                e = &HFA
    Case "ZOOM":                e = &HFB
    Case "NONAME":              e = &HFC
    Case "PA1":                 e = &HFD
    Case "OEM_CLEAR":           e = &HFE
    End Select
    EVirtualKeyCodes_Parse = e
End Function
' ^ ' ############################## ' ^ '    EVirtualKeyCodes     ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '     EKeyEventFlags      ' v ' ############################## ' v '
Public Function EKeyEventFlags_ToStr(ByVal e As EKeyEventFlags) As String
    Dim s As String, sOr As String: sOr = "Or "
    If e = 0 Then s = s & IIf(Len(s), sOr, "") & "KEYDOWN " 'hmm if its not keyup then its keydown, isn't it?
    If e And EKeyEventFlags.KEYEVENTF_EXTENDEDKEY Then s = s & IIf(Len(s), sOr, "") & "EXTENDEDKEY "
    If e And EKeyEventFlags.KEYEVENTF_KEYUP Then s = s & IIf(Len(s), sOr, "") & "KEYUP "
    If e And EKeyEventFlags.KEYEVENTF_UNICODE Then s = s & IIf(Len(s), sOr, "") & "UNICODE "
    If e And EKeyEventFlags.KEYEVENTF_SCANCODE Then s = s & IIf(Len(s), sOr, "") & "SCANCODE "
    EKeyEventFlags_ToStr = s
End Function

Public Function EKeyEventFlags_Parse(ByVal s As String) As EKeyEventFlags
    Dim e As EKeyEventFlags
    If InStr(s, "EXTENDEDKEY") Then e = e Or EKeyEventFlags.KEYEVENTF_EXTENDEDKEY
    If InStr(s, "KEYUP") Then e = e Or EKeyEventFlags.KEYEVENTF_KEYUP
    If InStr(s, "UNICODE") Then e = e Or EKeyEventFlags.KEYEVENTF_UNICODE
    If InStr(s, "SCANCODE") Then e = e Or EKeyEventFlags.KEYEVENTF_SCANCODE
    EKeyEventFlags_Parse = e
End Function

Public Sub EKeyEventFlags_ToList(aLst)
    aLst.Clear
    Dim i As Long, s As String
    s = EKeyEventFlags_ToStr(i): If Len(s) Then aLst.AddItem s
    For i = 0 To 4
        s = EKeyEventFlags_ToStr(2 ^ i): If Len(s) Then aLst.AddItem s
    Next
End Sub

Public Property Get ListBox_EKeyEventFlags(this As ListBox) As EKeyEventFlags
    'reads the selected elements in the Listbox into e
    Dim i As Long, s As String, tmpe As EKeyEventFlags
    Dim e As EKeyEventFlags
    For i = 0 To this.ListCount - 1
        If this.Selected(i) Then
            s = this.List(i)
            tmpe = EKeyEventFlags_Parse(s)
            e = e Or tmpe
        End If
    Next
    ListBox_EKeyEventFlags = e
End Property
Public Property Let ListBox_EKeyEventFlags(this As ListBox, ByVal e As EKeyEventFlags)
    'selects the elements in the Listbox if in e
    Dim i As Long, s As String, tmpe As EKeyEventFlags
    If e = 0 Then this.Selected(0) = True
    For i = 0 To this.ListCount - 1
        s = this.List(i)
        tmpe = EKeyEventFlags_Parse(s)
        If e And tmpe Then this.Selected(i) = True
    Next
End Property

'Public Function EKeyEventFlags_Read(aLst As ListBox) As EKeyEventFlags
'    'selects the elements in then Listbox if in e
'    Dim i As Long, s As String, tmpe As EKeyEventFlags
'    For i = 0 To aLst.ListCount - 1
'        s = aLst.List(i)
'        tmpe = EKeyEventFlags_Parse(s)
'        If e And tmpe Then aLst.Selected(i) = True
'    Next
'End Function

' ^ ' ############################## ' ^ '     EKeyEventFlags      ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    EMouseEventFlags     ' v ' ############################## ' v '

Public Function EMouseEventFlags_ToHex(ByVal e As EMouseEventFlags) As String
    Dim s As String: s = Hex(e)
    EMouseEventFlags_ToHex = "&&H" & s & "&&"
End Function

Public Function EMouseEventFlags_ToStr(ByVal e As EMouseEventFlags) As String
    Dim s As String, sOr As String: sOr = "Or "
    If e And EMouseEventFlags.MOUSEEVENTF_MOVE Then s = s & IIf(Len(s), sOr, "") & "MOVE "
    If e And EMouseEventFlags.MOUSEEVENTF_LEFTDOWN Then s = s & IIf(Len(s), sOr, "") & "LEFTDOWN "
    If e And EMouseEventFlags.MOUSEEVENTF_LEFTUP Then s = s & IIf(Len(s), sOr, "") & "LEFTUP "
    If e And EMouseEventFlags.MOUSEEVENTF_RIGHTDOWN Then s = s & IIf(Len(s), sOr, "") & "RIGHTDOWN "
    If e And EMouseEventFlags.MOUSEEVENTF_RIGHTUP Then s = s & IIf(Len(s), sOr, "") & "RIGHTUP "
    If e And EMouseEventFlags.MOUSEEVENTF_MIDDLEDOWN Then s = s & IIf(Len(s), sOr, "") & "MIDDLEDOWN "
    If e And EMouseEventFlags.MOUSEEVENTF_MIDDLEUP Then s = s & IIf(Len(s), sOr, "") & "MIDDLEUP "
    If e And EMouseEventFlags.MOUSEEVENTF_XDOWN Then s = s & IIf(Len(s), sOr, "") & "XDOWN "
    If e And EMouseEventFlags.MOUSEEVENTF_XUP Then s = s & IIf(Len(s), sOr, "") & "XUP "
    If e And EMouseEventFlags.MOUSEEVENTF_WHEEL Then s = s & IIf(Len(s), sOr, "") & "WHEEL "
    If e And EMouseEventFlags.MOUSEEVENTF_HWHEEL Then s = s & IIf(Len(s), sOr, "") & "HWHEEL "
    If e And EMouseEventFlags.MOUSEEVENTF_MOVE_NOCOALESCE Then s = s & IIf(Len(s), sOr, "") & "MOVE_NOCOALESCE "
    If e And EMouseEventFlags.MOUSEEVENTF_VIRTUALDESK Then s = s & IIf(Len(s), sOr, "") & "VIRTUALDESK "
    If e And EMouseEventFlags.MOUSEEVENTF_ABSOLUTE Then s = s & IIf(Len(s), sOr, "") & "ABSOLUTE "
    EMouseEventFlags_ToStr = s
End Function

Public Function EMouseEventFlags_Parse(ByVal s As String) As EMouseEventFlags
    Dim e As EMouseEventFlags
    If InStr(s, "MOVE ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_MOVE
    If InStr(s, "LEFTDOWN ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_LEFTDOWN
    If InStr(s, "LEFTUP ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_LEFTUP
    If InStr(s, "RIGHTDOWN ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_RIGHTDOWN
    If InStr(s, "RIGHTUP ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_RIGHTUP
    If InStr(s, "MIDDLEDOWN ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_MIDDLEDOWN
    If InStr(s, "MIDDLEUP ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_MIDDLEDOWN
    If InStr(s, "XDOWN ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_XDOWN
    If InStr(s, "XUP ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_XUP
    If InStr(s, "WHEEL ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_WHEEL
    If InStr(s, "HWHEEL ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_HWHEEL
    If InStr(s, "MOVE_NOCOALESCE ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_MOVE_NOCOALESCE
    If InStr(s, "VIRTUALDESK ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_VIRTUALDESK
    If InStr(s, "ABSOLUTE ") Then e = e Or EMouseEventFlags.MOUSEEVENTF_ABSOLUTE
    EMouseEventFlags_Parse = e
End Function

Public Sub EMouseEventFlags_ToList(aLst)
    aLst.Clear
    Dim i As Long, s As String
    's = EMouseEventFlags_ToStr(i): If Len(s) Then aLst.AddItem s '???
    For i = 0 To 15
        s = EMouseEventFlags_ToStr(2 ^ i): If Len(s) Then aLst.AddItem s
    Next
End Sub

Public Property Get ListBox_EMouseEventFlags(this As ListBox) As EMouseEventFlags
    'reads the selected elements in the Listbox into e
    Dim i As Long, s As String, tmpe As EMouseEventFlags
    Dim e As EKeyEventFlags
    For i = 0 To this.ListCount - 1
        If this.Selected(i) Then
            s = this.List(i)
            tmpe = EMouseEventFlags_Parse(s)
            e = e Or tmpe
        End If
    Next
    ListBox_EMouseEventFlags = e
End Property
Public Property Let ListBox_EMouseEventFlags(this As ListBox, ByVal e As EMouseEventFlags)
    'selects the elements in the Listbox if in e
    Dim i As Long, s As String, tmpe As EMouseEventFlags
    For i = 0 To this.ListCount - 1
        s = this.List(i)
        tmpe = EMouseEventFlags_Parse(s)
        If e And tmpe Then this.Selected(i) = True
    Next
End Property

Public Function ScreenCoords_ToMouseInpCoords(ByVal X_pix As Long, ByVal Y_pix As Long, MouseInp_X_out As Long, MouseInp_Y_out As Long) As Boolean
Try: On Error GoTo Catch
    Dim Res_W As Double: Res_W = Screen.Width / Screen.TwipsPerPixelX  ' 2560
    Dim Res_H As Double: Res_H = Screen.Height / Screen.TwipsPerPixelY ' 1440
    MouseInp_X_out = X_pix * 65535# / Res_W
    MouseInp_Y_out = Y_pix * 65535# / Res_H
    'Debug.Print MouseInp_X_out & "; " & MouseInp_Y_out
    ScreenCoords_ToMouseInpCoords = True
    Exit Function
Catch:
End Function

Public Function MouseInpCoords_ToScreenCoords(ByVal MouseInp_X As Long, ByVal MouseInp_Y As Long, X_pix_out As Long, Y_pix_out As Long)
Try: On Error GoTo Catch
    Dim Res_W As Double: Res_W = Screen.Width / Screen.TwipsPerPixelX  ' 2560
    Dim Res_H As Double: Res_H = Screen.Height / Screen.TwipsPerPixelY ' 1440
    X_pix_out = MouseInp_X * Res_W / 65535#
    Y_pix_out = MouseInp_Y * Res_H / 65535#
    'Debug.Print X_pix_out & "; " & Y_pix_out
    MouseInpCoords_ToScreenCoords = True
    Exit Function
Catch:
End Function
' ^ ' ############################## ' ^ '    EMouseEventFlags     ' ^ ' ############################## ' ^ '

