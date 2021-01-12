local VBS = require "VBS"

local VBF = require "VBF"

VBS.Init()
    VBF.InputBox('a', 'Type your text', 'World gonna say:')
    VBF.MsgBox('x', '', VBS.Buttons.OkOnly, 'World Says:', '& a') --Even if you gonna concat just one var you need  do this (i need fucking get real)
VBS.End()

print('Ya like jazz?')


VBS.Run()