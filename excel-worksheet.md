# Tokyo Midnight Theme for Excel (VBA Macro)

This repository provides a simple VBA macro to apply a **Tokyo Midnight** dark theme to any Excel worksheet, along with a one-click **Undo** function that restores the original formatting using a hidden backup sheet.

Inspired by the sleek aesthetics of dark IDEs and modern dashboards, the theme enhances visibility and provides a stylish night-mode UI experience.

---

## Features

- Applies Tokyo Midnight colors to the active worksheet
- Sets fonts to `Segoe UI` with lighter text on a dark background
- Highlights the header row with a blue accent
- Adds subtle cell borders for clean separation
- Creates a hidden backup sheet (`Backup`) before making changes
- One-click restore to original formatting

---

## Installation

1. Press `ALT + F11` to open the **VBA Editor** in Excel.
2. Insert a new **Module** (`Insert > Module`).
3. Copy and paste the following code:

```vba
Sub TokyoMidnightStyleWithBackup()
    Dim ws As Worksheet
    Dim backupWs As Worksheet

    On Error Resume Next
    Set backupWs = Sheets("Backup")
    If backupWs Is Nothing Then
        Set backupWs = Sheets.Add
        backupWs.Name = "Backup"
        backupWs.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    ActiveSheet.Cells.Copy Destination:=backupWs.Cells

    Set ws = ActiveSheet
    With ws.Cells
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Interior.Color = RGB(20, 20, 30)
        .Font.Color = RGB(230, 230, 230)
    End With

    With ws.Rows(1)
        .Interior.Color = RGB(0, 120, 215)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.UsedRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(60, 60, 70)
        .Weight = xlThin
    End With
End Sub

Sub UndoTokyoMidnightStyle()
    Dim ws As Worksheet
    Dim backupWs As Worksheet

    Set ws = ActiveSheet
    Set backupWs = Sheets("Backup")

    If Not backupWs Is Nothing Then
        backupWs.Cells.Copy Destination:=ws.Cells
    Else
        MsgBox "No backup available to restore.", vbExclamation
    End If
End Sub
````

---

## Usage

1. Select any worksheet.
2. Run `TokyoMidnightStyleWithBackup` to apply the dark theme.
3. If needed, run `UndoTokyoMidnightStyle` to restore original formatting.

> **Note:** The macro creates a backup copy before applying the style and hides it as a `VeryHidden` sheet, preventing accidental edits.

---

## Customization

You can tweak:

* Font style and size
* Accent and background colors
* Header formatting
* Border visibility and weight

---

## License

MIT License — use, modify, and distribute freely with attribution.

---

Maintained by [@damian-dev1](https://github.com/damian-dev1)
