local VBF = {}

local VBS = require "VBS"

function VBF.MsgBox(varname, Text, config, title, ...)
    if ... == nil then
        assert(VBS.file:write(varname..'=MsgBox("'..Text..'", '..config..', "'..title..'")\n'), "Cannot open VBS, check if a proccess its using it")
    else
        assert(VBS.file:write(varname..'=MsgBox("'..Text..'"'..(...)..', '..config..', "'..title..'")\n'), "Cannot open VBS, check if a proccess its using it")
    end
end

function VBF.InputBox(varname, text, title, ...)
    if ... == nil then
        assert(VBS.file:write(varname..'=InputBox("'..text..'", "'..title..'")\n'), "Cannot open VBS, check if a proccess its using it")
    else
        assert(VBS.file:write(varname..'=InputBox("'..text..'"'..(...)..', "'..title..'")\n'), "Cannot open VBS, check if a proccess its using it")
    end
end

function VBF.CDate(varname, date)
    assert(VBS.file:write(varname..'=CDate"("'..date..'")\n'), "Cannot open VBS, check if a proccess its using it")
end

function VBF.Date(varname)
    assert(VBS.file:write(varname..'=Date\n'),  "Cannot open VBS, check if a proccess its using it")
end

function VBF.DateAdd(varname, interval, number, date)
    assert(VBS.file:write(varname..'=DateAdd("'..interval..'", "' ..number..'", "'..date..'")\n'), "Cannot open VBS, check if a proccess its using it")
end

function VBF.Array(varname, ...)
    assert(VBS.file:write(varname.."=Array("..(...)..")", "Cannot open VBS, check if a proccess its using it"))
end

function VBF.Split(varname, split)
    assert(VBS.file:write(varname.."=Split("..split..")", "Cannot open VBS, check if a proccess its using it"))
end

function VBF.CreateObject(varname, Object)
    assert(VBS.file:write(varname.."=CreateObject("..Object..")", "Cannot open VBS, check if a proccess its using it"))
end

return VBF