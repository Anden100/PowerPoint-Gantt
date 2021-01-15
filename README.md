# PowerPoint-Gantt
Simple VBA class to create aesthetically pleasing and customizable Gantt Charts in PowerPoint

# Usage example
```vb
Sub run()
    Dim Gantt As New GanttChart

    Dim ColorSecondary, ColorGold As Long

    ColorSecondary = RGB(0, 0, 255)
    ColorGold = RGB(255, 215, 0)

    Gantt.MarginTop = 122

    Gantt.OutlineColor = RGB(89, 89, 89)
    Gantt.LabelFontColor = RGB(13, 13, 13)
    Gantt.HeaderFontColor = RGB(38, 38, 38)

    Gantt.HeaderFontSize = 8
    Gantt.LabelFontSize = 10

    Gantt.StartDate = DateSerial(2021, 1, 1)
    Gantt.EndDate = DateSerial(2021, 6, 30)

    Gantt.ResponsibleColumn = False

    Gantt.AddActivity "Project Design", DateSerial(2021, 1, 1), DateSerial(2021, 2, 15), "", ColorSecondary
    Gantt.AddActivity "Model review 1", DateSerial(2021, 2, 7), DateSerial(2021, 2, 22), "", ColorSecondary
    Gantt.AddActivity "System development", DateSerial(2021, 2, 10), DateSerial(2021, 3, 15), "", ColorSecondary

    Gantt.AddActivity "Rollout phase 1", DateSerial(2021, 3, 15), DateSerial(2021, 3, 31), "", ColorGold

    Gantt.AddActivity "Learnings", DateSerial(2021, 3, 1), DateSerial(2021, 4, 15), "", ColorSecondary
    Gantt.AddActivity "Adjustments", DateSerial(2021, 4, 7), DateSerial(2021, 4, 22), "", ColorSecondary
    Gantt.AddActivity "Model review 2", DateSerial(2021, 4, 10), DateSerial(2021, 5, 15), "", ColorSecondary

    Gantt.AddActivity "Full scale rollout", DateSerial(2021, 5, 15), DateSerial(2021, 5, 31), "", ColorGold

    Gantt.AddCallout "Deadline 1", DateSerial(2021, 3, 15)
    Gantt.AddCallout "Deadline 2", DateSerial(2021, 5, 15)

    Gantt.Generate
End Sub
```
