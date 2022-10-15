# PropDumper - File Property/Metadata Dumper
![Screenshot](https://i.imgur.com/jYWgtYA.jpg)

This program will create a dump of all properties and metadata visible to Explorer which aren't blank for the given file.

You can create these output files with the specified extension, in the same folder where the file resides or a central folder. If the output already exists, there's an option to rename to the next available name the same way Windows does, File (1).txt, File (2).txt, whatever number is available.

Properties include hidden Unicode codepoints indicating whether it's left to right or right to left; I usually view text in Notepad with Western script instead of Unicode, so I find it useful to remove these. So that option is there.

This will not go into zip/cab files, even though Windows considers them folders now.

This is all done by using the Windows Property System (IPropertyStore et al), which is the system that underlies how all the properties are displayed in Explorer-- so you're able to list exactly what Explorer is, without having to worry about having to derive things like image width/height yourself. If Explorer can see it, it will be in the dump. Note: Some properties like Office are derived through 64bit shell extensions that won't load into 32bit apps; so if you want to make sure it's an identical representation, use the 64bit build for 64bit Windows.

This program was created in [twinBASIC](https://github.com/twinbasic/twinbasic), an actual successor to VB6/VBA. It's very far along at this point; language compatibility is nearly complete, and basic Forms with most of the default controls can be created, ActiveX control support is decent (but can't use .ctl controls yet). It can compile 64bit exes using VBA7/64 syntax from Office, and in addition there's a bunch of new features bringing the language into this century. You'll see the splash screen for it in the x64 build as I don't currently have a paid subscription (and that splash screen is currently the only limitation of free version). 

One of the new features is defining COM interfaces locally; I made this project primarily to test things in my [tbShellLib library](https://github.com/fafalone/tbShellLib), a collection of Windows shell interfaces and other COM components that's a 64bit compatible successor my [oleexp.tlb](https://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb) project for VB6. Note that this is why the source code file is so large; it imports the entire library. Which is quite expansive.

(Note that since this project uses new language features in tB, it would take some work to backport it to VB6-- tB is backwards compatible, not forwards compatible, just like VBA7x64 isn't compatible with 32bit VBA6 or VB6). 

## Requirements
Windows 7+

Source and binaries are self-contained, no installation or additional dependencies.

Building from source requires [twinBASIC Beta 147](https://github.com/twinbasic/twinbasic/releases) or newer.

## SOURCE CODE NOTES

For anyone unfamiliar with twinBASIC source code structure, it combines all source into a single file. The .twinproj file in the root above is the complete source code; it's large because it includes the library dependencies- tbShellLib (my project) and the twinBASIC WinNativeForms package that create the Forms/Controls, as well as the program icon and manifest. You can browse the local project source code to this program itself in the SourceBrowse folder, but you need the .twinproj file to open and compile it yourself in twinBASIC. You can also browse the source for tbShellLib [in it's repository](https://github.com/fafalone/tbShellLib).

## Dumping properties

Windows shell interfaces and Property System make dumping properties very easy:

```
Private Sub DumpFileProperties(siFile As IShellItem, sName As String)
    Dim sOut As String
    Dim lpPath As LongPtr, sPath As String
    siFile.GetDisplayName SIGDN_FILESYSPATH, lpPath
    sPath = LPWSTRtoStr(lpPath)

    
    sOut = "Filename: " & sName & vbCrLf
    sOut &= "File full path: " & sPath & vbCrLf

    Dim si2 As IShellItem2
    Dim pps As IPropertyStore
    Dim ppd As IPropertyDescription
    Dim lpFmt As LongPtr, sPropFmt As String
    Dim sFileOut As String
    
    If bCurDir Then
        sFileOut = sPath & sExt
    Else
        sFileOut = sPathOut & sExt
    End If
    If PathFileExistsW(sFileOut) Then
        If chkRename.Value = vbChecked Then
            sFileOut = UniqueNameInSeq(sFileOut)
        Else
            AppendLog "Skipping " & sName & "; output file exists."
            Exit Sub
        End If
    End If
    
    Set si2 = siFile
    si2.GetPropertyStore GPS_DEFAULT Or GPS_BESTEFFORT Or GPS_OPENSLOWITEM, IID_IPropertyStore, pps
    If (pps Is Nothing) = False Then
        Dim nMax As Long
        pps.GetCount nMax
        If nMax Then
        	AppendLog "Dumping " & nMax & " properties for " & sName & "..."
                
            Dim i As Long
            Dim pk As PROPERTYKEY
            Dim lpProp As LongPtr, sProp As String
            Dim lpPropC As LongPtr, sPropC As String
            Dim lpPropN As LongPtr, sPropN As String
            For i = 0 To nMax - 1
                pps.GetAt i, pk
                If pk.fmtid.Data1 <> 0 Then
                    PSGetPropertyDescription pk, IID_IPropertyDescription, ppd
                    If (ppd Is Nothing) = False Then
                        ppd.GetDisplayName lpPropN
                        sPropN = LPWSTRtoStr(lpPropN)
                        ppd.GetCanonicalName lpPropC
                        sPropC = LPWSTRtoStr(lpPropC)
                        PSFormatPropertyValue ObjPtr(pps), ObjPtr(ppd), PDFF_DEFAULT, lpProp
                        sProp = LPWSTRtoStr(lpProp)
                        If bRemoveFmt Then RemoveFormatChars(sProp)
                        If sProp <> "" Then sOut &= sPropN & " (" & sPropC & ")=" & sProp & vbCrLf
                    Else
                        Debug.Print "Couldn't get propdesc for " & dbg_PKEYToString(pk)
                    End If
                End If
            Next i
        Else
            sOut &= "No properties listed for file."
        End If
    Else
        AppendLog "Couldn't open property store for " & sPath
        sOut &= "Couldn't open property store for this file."
    End If
    

    WriteStrToFile sOut, sFileOut
    nCount = nCount + 1
End Sub
```
