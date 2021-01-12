# Using

1- Require VBS and VBF
```lua
    local VBS = require "VBS"
    local VBF = require "VBF"
```

2- Use VBS.Init to open file and VBS.End to close
```lua
    VBS.Init()
        --Its an section so press TAB to your code looks pretty
    VBS.End()
```

# Hello World!
    If you are a newgen to do an hello world program in vbs,

    You need store a MsgBox into a variable

    Using VBF in the VBS.Init Section, we have 5 params

    Variable (Gonna store the message box)

    Text (Where the text appears)

    Config (Add buttons and icons, you need put VBS.Buttons.OkOnly to default)

    Title (Title of the GUI)

    VarArg (Concat things in Text, you can make it nil and the script solves that)

    ```lua
        VBF.MsgBox('x', 'Hello!', VBS.Buttons.OKOnly, 'World Says:')
    ```

    Now lets run, use VBS.Run outside the VBS.Init and VBS.End

    Should a Gui appears saying hello

    Doc Folder Coming soon bro :>