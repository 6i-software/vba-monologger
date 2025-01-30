---
description: Understand how to use channels in VBA Monologger to effectively categorize and manage log messages.
---

## What is a channel?

A [channel](../introduction.md#identifying-a-logger-with-a-channel) refers to the name given to a specific logger.

This name helps categorize and manage log messages effectively. By using different channels, each logger can be configured with specific handlers, formatters, and log levels. This ensures that logs from different parts of an application are handled appropriately and organized for easy analysis. 

It is a powerful way to identify which part of an application a log entry is associated with. This is especially useful in large applications with multiple components and multiple loggers.


## Set log channel into built-in loggers

In each default built-in loggers provided by the VBA Monologger factory (*e.g.* `LoggerConsoleVBA`, `LoggerConsole` or `LoggerFile`) you can set the name with its first parameter `paramLoggerName`.

```vbscript
Public Sub howto_set_logger_name()
    ' Create a logger instance with the channel name "App"
    Dim Logger As VBAMonologger.LoggerInterface
    Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA("App")
    
    ' Same usage for another default loggers
    ' Set Logger = VBAMonologger.Factory.createLoggerConsole("App")
    ' Set Logger = VBAMonologger.Factory.createLoggerFile("App")  
 
    ' Use the logger for each severity levels
    Logger.trace "Authentication function call for user 'Bob Morane'." 
    Logger.info "User 'UltraVomit' has logged in successfully."
    Logger.notice "Process completed successfully with minor issues."
    Logger.warning "'Beetlejuice' should not be called more than 3 times."
    Logger.error "An error occurred with the user 'DRZCFOS2'."
    Logger.critical "System is in an unstable state."
    Logger.alert "Action required: unable to generate the dashboard."
    Logger.emergency "A critical failure occurred in the application."
End Sub
```

``` title='Result in VBA console'
[2024/12/16 12:51:17] App.DEBUG: Authentication function call for user 'Bob Morane'.
[2024/12/16 12:51:17] App.INFO: User 'UltraVomit' has logged in successfully.
[2024/12/16 12:51:17] App.NOTICE: Process completed successfully with minor issues.
[2024/12/16 12:51:17] App.WARNING: The user 'Beetlejuice' should not be called more than 3 times.
[2024/12/16 12:51:17] App.ERROR: An error occurred when the user 'DeadRobotZombieCopFromOuterspace' tried to read the dashboard file.
[2024/12/16 12:51:17] App.CRITICAL: System is in an unstable state. Unable to authenticate the user 'Skjalg Skagen'.
[2024/12/16 12:51:17] App.ALERT: Action required: unable to generate the dashboard.
[2024/12/16 12:51:17] App.EMERGENCY: A critical failure occurred in the application for moving files.
```


## Set logger's name into empty logger

If you create a simple logger without handlers and pre-processors, as an empty logger, use the `name` property to change its channel name.

```vbscript
Public Sub howto_set_logger_name()
    ' Create an emtpy logger (without handler, pre-processors...)
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLogger()
    
    ' Set the channel name
    Logger.name = "App"
End Sub
```



