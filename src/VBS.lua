local VBS = {}

VBS._VERSION = "1.1 BETA"

VBS.Buttons = {
    OkOnly = "vbOkOnly";
    Ok = "vbOk"
}

VBS.Icons = {
    Critical = "vbCritical";
}

VBS.Controls = {
    NewLine = 'vbNewLine'
}

VBS.MainFile = "Executor.vbs" --Default

function VBS.Init()
    VBS.file = io.open(VBS.MainFile, 'w+')

    function VBS.Var(typ, name, value)
        if value == nil then
            if typ == "Dim" then
                assert(VBS.file:write("Dim "..name.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Static" then
                assert(VBS.file:write("Static "..name.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Public" then
                assert(VBS.file:write("Public "..name.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Private" then
                assert(VBS.file:write("Private "..name.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Set" then
                error("Set cannot be used while the variable has no value!")
            elseif typ == "None" then
                error("None cannot be used while the variable has no value!")
            else
                error("Variable Type '"..typ.."' dont exist.")
            end
        else
            if typ == "Dim" then
                assert(VBS.file:write("Dim "..name.."="..value.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Static" then
                assert(VBS.file:write("Static "..name.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Public" then
                assert(VBS.file:write("Public "..name.."="..value.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Private" then
                assert(VBS.file:write("Private "..name.."="..value.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Set" then
                assert(VBS.file:write("Set "..name.."="..value.."\n"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "None" then
                assert(VBS.file:write(name.."="..value.."\n"), "Cannot open VBS, check if a proccess its using it")
            else
                error("Variable Type '"..typ.."' dont exist.")
            end
        end
    end

    function VBS.VarAs(typ, name, value)
        if typ == "Dim" then
            assert(VBS.file:write("Dim "..name.." As "..value.."\n"), "Cannot open VBS, check if a proccess its using it")
        elseif typ == "Static" then
            assert(VBS.file:write("Dim "..name.." As "..value.."\n"), "Cannot open VBS, check if a proccess its using it")
        elseif typ == "Public" then
            assert(VBS.file:write("Public "..name.." As "..value.."\n"), "Cannot open VBS, check if a proccess its using it")
        elseif typ == "Private" then
            assert(VBS.file:write("Private "..name.." As "..value.."\n"), "Cannot open VBS, check if a proccess its using it")
        elseif typ == "Set" then
            assert(VBS.file:write("Set "..name.." As "..value.."\n"), "Cannot open VBS, check if a proccess its using it")
        elseif typ == "None" then
            assert(VBS.file:write(name.." As "..value.."\n"), "Cannot open VBS, check if a proccess its using it")
        else
            error("Variable Type '"..typ.."' dont exist.")
        end
    end

    function VBS.If(condition)
        assert(VBS.file:write("If "..condition.." Then\n\t"), "Cannot open VBS, check if a proccess its using it")
        
        function VBS.Else()
            VBS.file:write('Else\n\t')
        end
    
        function VBS.ElseIf(cond)
            VBS.file:write('Elseif '..cond..' Then\n\t')
        end
    
        function VBS.EndIf()
            VBS.file:write('End If\n')
        end
    end
    
    function VBS.Function(typ, name, ...)
        if ... == nil then
            if typ == "Static" then
                assert(VBS.file:write("Static Function "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Public" then
                assert(VBS.file:write("Public Function "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Private" then
                assert(VBS.file:write("Private Function "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "None" then
                assert(VBS.file:write("Function "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            else
                error("Function Type '"..typ.."' dont exist.")
            end
        else
            if typ == "Static" then
                assert(VBS.file:write("Static Function "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Public" then
                assert(VBS.file:write("Public Function "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Private" then
                assert(VBS.file:write("Private Function "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "None" then
                assert(VBS.file:write("Function "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            else
                error("Function Type '"..typ.."' dont exist.")
            end
        end
    
        function VBS.EndFunction()
            VBS.file:write('End Function\n')
        end
    end
    
    function VBS.For(var, value, count)
        assert(VBS.file:write("For "..var.." = "..value.." to "..count.."\n\t"))
        function VBS.ExitFor()
            assert(VBS.file:write('Exit For\n'))
        end
        function VBS.Next()
            VBS.file:write("Next\n")
        end
    end
    
    function VBS.ForEach(var, item)
        assert(VBS.file:write("For Each "..var.." in "..item.."\n\t"), "Cannot open VBS, check if a proccess its using it")
        function VBS.ExitFor()
            assert(VBS.file:write('Exit For\n'))
        end
        function VBS.Next()
            VBS.file:write("Next\n")
        end
    end
    
    function VBS.While(condition)
        assert(VBS.file:write("While "..condition.."\n\t"), "Cannot open VBS, check if a proccess its using it")
        function VBS.Wend()
            VBS.file:write("\nWend\n")
        end
    end
    
    function VBS.DoWhile(condition)
        assert(VBS.file:write("Do While "..condition.."\n\t"), "Cannot open VBS, check if a proccess its using it")
        function VBS.ExitDo()
            assert(VBS.file:write('Exit Do\n'))
        end
        function VBS.Loop()
            VBS.file:write('Loop')
        end
    end
    
    function VBS.DoUntil(condition)
        assert(VBS.file:write("Do Until "..condition.."\n\t"), "Cannot open VBS, check if a proccess its using it")
        function VBS.ExitDo()
            assert(VBS.file:write('Exit Do\n'))
        end
        function VBS.Loop()
            VBS.file:write('Loop')
        end
    end

    
    function VBS.Call(func, ...)
        if ... == nil then
            assert(VBS.file:write("Call "..func.."()"), "Cannot open VBS, check if a proccess its using it")
        else
            assert(VBS.file:write("Call "..func.."("..(...)..")"), "Cannot open VBS, check if a proccess its using it")
        end
    end
    
    function VBS.DoLoop()
        assert(VBS.file:write("Do\n\t"), "Cannot open VBS, check if a proccess its using it")
        function VBS.ExitDo()
            VBS.file:write('Exit Do\n')
        end
        function VBS.Loop()
            VBS.file:write('Loop')
        end
    end

    function VBS.Select(var)
        assert(VBS.file:write("Select case "..var.."\n\t"), "Cannot open VBS, check if a proccess its using it")
        function VBS.Case(case)
            VBS.file:write('case '..case)
        end
        function VBS.EndSelect()
            VBS.file:write('End Select')
        end
    end

    function VBS.Sub(typ, name, ...)
        if ... == nil then
            if typ == "Static" then
                assert(VBS.file:write("Static Sub "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Public" then
                assert(VBS.file:write("Public Sub "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Private" then
                assert(VBS.file:write("Private Sub "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "None" then
                assert(VBS.file:write("Sub "..name.."()\n\t"), "Cannot open VBS, check if a proccess its using it")
            else
                error("Function Type '"..typ.."' dont exist.")
            end
        else
            if typ == "Static" then
                assert(VBS.file:write("Static Sub "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Public" then
                assert(VBS.file:write("Public Sub "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "Private" then
                assert(VBS.file:write("Private Sub "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            elseif typ == "None" then
                assert(VBS.file:write("Sub "..name.."("..(...)..")\n\t"), "Cannot open VBS, check if a proccess its using it")
            else
                error("Sub Type '"..typ.."' dont exist.")
            end
        end
    end

    function VBS.Class(name)
        assert(VBS.file:write("Class "..name.."\n\t"), "Cannot open VBS, check if a proccess its using it")
        function VBS.Property(typ, kyw, name, ...)
            if ... == nil then
                if typ == "Static" then
                    VBS.file:write("Static Property "..kyw.." "..name.."()\n\t")
                elseif typ == "Public" then
                    VBS.file:write("Public Property "..kyw.." "..name.."()\n\t")
                elseif typ == "Private" then
                    VBS.file:write("Private Property "..kyw.." "..name.."()\n\t")
                elseif typ == "None" then
                    VBS.file:write("Property "..kyw.." "..name.."()\n\t")
                else
                    error("Property Type '"..typ.."' dont exist.")
                end
            else
                if typ == "Static" then
                    VBS.file:write("Static Property "..kyw.." "..name.."("..(...)..")\n\t")
                elseif typ == "Public" then
                    VBS.file:write("Public Property "..kyw.." "..name.."("..(...)..")\n\t")
                elseif typ == "Private" then
                    VBS.file:write("Private Property "..kyw.." "..name.."("..(...)..")\n\t")
                elseif typ == "None" then
                    VBS.file:write("Property "..kyw.." "..name.."("..(...)..")\n\t")
                else
                    error("Property Type '"..typ.."' dont exist.")
                end
            end
        end
        function VBS.EndClass()
            VBS.file:write('End Class')
        end
    end

    function VBS.End()
        VBS.file:close()
    end
end

VBS.Run = function()
    os.execute('start Executor.vbs')
end

return VBS