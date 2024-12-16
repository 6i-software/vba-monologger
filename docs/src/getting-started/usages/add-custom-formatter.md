## What is a log formatter?

A [log handler](../introduction.html#processing-log-records-with-a-handler) is a key component responsible for processing each log entry. When a log message is generated, it is not simply recorded; it must be directed to a location where it can be viewed and used. This is where the handler comes in, determining where and how each log entry will be sent or saved (show into console, send to a file, send by mail...). And each "log handler" is associated with a unique [log formatter](../introduction.html#formatting-log-records-the-serialization-of-logs-record). 

The formatter is a specialized component responsible for defining the structure and presentation of log messages. The formatter processes and organizes each log entry, converting it from its raw form to a readable format tailored to a specific type (text, HTML, JSON, etc.). This process can be seen as similar to the serialization of a log record.

## Default formatter used in built-in loggers

When creating loggers with the factory methods provided by `VBAMonologger.Factory`, the type of handler and its formatter depend on the target output (VBA console, file, or Windows console). 

| **Factory method**         | **Default handler** | **Default formatter**                                                                                                              |
|----------------------------|---------------------|------------------------------------------------------------------------------------------------------------------------------------|
| `createLoggerConsoleVBA()` | `HandlerConsoleVBA` | `FormatterLine`                                                                                                                    |
| `createLoggerFile()` | `HandlerFile`       | `FormatterLine`                                                                                                                    |
| `createLoggerConsole()` | `HandlerConsole`    | `FormatterANSIcoloredLine` if ANSI color support is enabled with `paramWithANSIColorSupport=true`), and otherwise `FormatterLine`. |

As you can see in each factory method's signatures, you can use the option `paramFormatter` to load a custom formatter.

=== "createLoggerConsoleVBA"

    ``` vbscript
    Public Function createLoggerConsoleVBA( _
        Optional ByVal paramLoggerName As String = vbNullString, _
        Optional ByRef paramFormatter As FormatterInterface = Nothing _
    ) As VBAMonologger.Logger
    ```

=== "createLoggerFile"

    ``` vbscript
    Public Function createLoggerFile( _
        Optional ByVal paramLoggerName As String = vbNullString, _
        Optional ByRef paramFormatter As FormatterInterface = Nothing, _
        Optional ByVal paramLogFileName As String = vbNullString, _
        Optional ByVal paramLogFileFolder As String = vbNullString _
    ) As VBAMonologger.Logger
    ```

=== "createLoggerConsole"

    ``` vbscript
    Public Function createLoggerConsole( _
        Optional ByVal paramLoggerName As String = vbNullString, _
        Optional ByRef paramFormatter As FormatterInterface = Nothing, _
        Optional ByRef paramWithANSIColorSupport As Boolean = True, _
        Optional ByRef paramWithNewlineForContextAndExtra As Boolean = True, _
        Optional ByRef paramWithDebugServer As Boolean = False, _
        Optional ByRef paramWithDebugClient As Boolean = False _
    ) As VBAMonologger.Logger
    ```


## Create a custom formatter

To illustrate how to load a custom formatter, we will create a new line formatter with a different line template, but it works the same way with any formatter.

The `FormatterLine` uses a **line template**, as a string, in order to format each log entry. This line template defines the representation of a log entry, with **placeholders** that will be replaced by actual values from the log record. So you can customize the template to fit your needs. For example, if you prefer a simpler format or want to add additional information, you can easily adjust the template using the `templateLine` property.

> **Understanding the line template behavior**
>
> The formatting system of `FormatterLine` uses regular expressions to handle placeholders within the line template, allowing them to be defined with *prefixes* and *suffixes* to modify their final output. The prefix text is added before the placeholder’s value, and the suffix text is added after the placeholder's value. And if a placeholder has no value, then the prefix and postfix are not displayed in the final output.
>
> Here’s how a placeholder can be structured in the template:
>
> ``` twig
> {{ <prefix/> placeholder <suffix/> }}
> ```
>
> By default, the line template of `VBAMonologger.Formatter.FormatterLine` looks like the following:
>
> ``` twig title="Line template"
> [{{datetime}}] {{channel}}.{{level_name}}: {{message}}{{< | ctx=/> context}}{{< | extra=/> extra}}
> ```
>
> The following placeholders are included:
>
> | Placeholder&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description                                                                                                                                    |
> |-------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------|
> | `{{datetime}}`                                                                            | The date and time of the log entry.                                                                                                            |
> | `{{channel}}`                                                                             | The channel (or source) from which the log originates.                                                                                         |
> | `{{level_name}}`                                                                          | The log level (*e.g.*, `INFO`, `ERROR`, `DEBUG` ...).                                                                                          |
> | `{{message}}`                                                                             | The main log message.                                                                                                                          |
> | `{{extra}}`                                                                               | Extra metadata or custom information attached to the log entry. This can include arbitrary key-value pairs, typically added by pre-processors. |

Here's an example of a new custom line formatter.

```vbscript 
Public Sub howto_create_custom_formatter()
    ' Create a custom formatter
    Dim customFormatter As VBAMonologger.FormatterLine
    Set customFormatter = VBAMonologger.Factory.createFormatterLine
    customFormatter.templateLine = ":: {{ channel }}.{{ level_name }} - {{ message }}"
End Sub
```

## Load a custom formatter into logger

Just fill the parameter `paramFormatter` in factory's method, or if you have an instance of `Handler` used by the logger, you can use the property `formatter` directly.

```vbscript 
Public Sub howto_change_formatter()
    ' Create a custom formatter
    Dim customFormatter As VBAMonologger.FormatterLine
    Set customFormatter = VBAMonologger.Factory.createFormatterLine
    customFormatter.templateLine = ":: {{ channel }}.{{ level_name }} - {{ message }}"
    
    ' Load custom formatter into logger
    Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA( _
        paramLoggerName:="App", _
        paramFormatter:=customFormatter _
    )
    
    ' Logs message for each severity levels
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

Result in VBA console:

=== "With custom formatter"

    ```
    :: App.DEBUG - Authentication function call for user 'Bob Morane'.
    :: App.INFO - User 'UltraVomit' has logged in successfully.
    :: App.NOTICE - Process completed successfully with minor issues.
    :: App.WARNING - The user 'Beetlejuice' should not be called more than 3 times.
    :: App.ERROR - An error occurred when the user 'DeadRobotZombieCopFromOuterspace' tried to read the dashboard file.
    :: App.CRITICAL - System is in an unstable state. Unable to authenticate the user 'Skjalg Skagen'.
    :: App.ALERT - Action required: unable to generate the dashboard.
    :: App.EMERGENCY - A critical failure occurred in the application for moving files.
    ```

=== "With default line formatter"

    ```
    [2024/12/16 12:51:17] App.DEBUG: Authentication function call for user 'Bob Morane'.
    [2024/12/16 12:51:17] App.INFO: User 'UltraVomit' has logged in successfully.
    [2024/12/16 12:51:17] App.NOTICE: Process completed successfully with minor issues.
    [2024/12/16 12:51:17] App.WARNING: The user 'Beetlejuice' should not be called more than 3 times.
    [2024/12/16 12:51:17] App.ERROR: An error occurred when the user 'DeadRobotZombieCopFromOuterspace' tried to read the dashboard file.
    [2024/12/16 12:51:17] App.CRITICAL: System is in an unstable state. Unable to authenticate the user 'Skjalg Skagen'.
    [2024/12/16 12:51:17] App.ALERT: Action required: unable to generate the dashboard.
    [2024/12/16 12:51:17] App.EMERGENCY: A critical failure occurred in the application for moving files.
    ```
