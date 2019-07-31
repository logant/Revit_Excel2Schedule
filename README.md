# Revit Excel to Revit Plugin

## Description
This is a plugin for Autodesk Revit that imports formatted excel content into an Excel schedule. It can import or link, allowing for updates to happen when the Revit file opens, or through a Manage Excel Links command.

## Dependencies
Uses [RevitCommon](https://github.com/logant/RevitCommon) for Revit UI integration (adding buttons to launch the commands) and for the Excel reading.

## Known Issues
I've noticed that cell boundaries sometimes don't come through for the overall outer boundary of the schedule.