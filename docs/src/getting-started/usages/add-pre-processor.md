## What is a log pre-processor?

[Pre-processors](../introduction.html#pre-processor-of-log-records) are a powerful feature, allowing for additional metadatas to be added to log messages before they are recorded. These functions can be used to enrich log messages with extra information that might not be directly part of the log entry itself, but is still relevant for better understanding and tracking. 

Pre-processors can modify, format, or even generate additional metadata that can be attached to each log message into the extra property. A logger can reference one or more pre-processors.


## Default pre-processor loaded in built-in loggers

When creating loggers with the factory methods provided by `VBAMonologger.Factory`, it automatically loads the pre-processor `ProcessorPlaceholders` for each built-in loggers. This ensures that all log entries include the placeholder's features.

| **Factory method**         | **Default pre-processor** |  
|----------------------------|---------------------------|  
| `createLoggerConsoleVBA()` | `ProcessorPlaceholders`   |  
| `createLoggerFile()`       | `ProcessorPlaceholders`   |  
| `createLoggerConsole()`    | `ProcessorPlaceholders`   |  

The `ProcessorPlaceholders`, allows to replace specific variables, or [placeholders](../introduction.html#adding-metadatas-in-log-records), in log messages with their corresponding values, adding dynamic context to the logs. It consumes the log context variable given with the log message.


## Available pre-processors

In VBA Monologger, several built-in processors offer specific functionalities to enhance the log entry by adding additional context or modifying it in various ways before it is passed to the handlers. Here are some examples of the available pre-processors:

| Pre-processor          | Description                                                                                                                                                                                                                                                                                               |
|-------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `ProcessorPlaceholders` | Allows to replace specific variables or placeholders in log messages with their corresponding values, adding dynamic context to the logs. The `context` variable must be a VBA `Scripting.Dictionary`. (*e.g.* `logger.info("Authentication failed for user '{UserName}' with id '{UserID}'.", context)`. |
| `ProcessorTags`         | Adds one or more tags to a log entry.                                                                                                                                                                                                                                                                     |
| `ProcessorUID`          | Adds a unique identifier (UID) to each session . The generated UID consists of hexadecimal characters, and its size can be configured.                                                                                                                                                                    |
| `ProcessorUsageMemory`  | Adds the computer's memory usage to each log entry. The system's current memory status is retrieved using the `GlobalMemoryStatusEx` API in Windows.                                                                                                                                                      |
| `ProcessorUsageCPU`     | Adds the computer's cpu usage to each log entry.                                                                                                                                                                                                                                                          |

Result of a logger with pre-processors : placeholders, tags (`environment, user_role`), UID, usage memory and usage CPU.

```
[2024/12/13 18:59:30] App.INFO: User '35353' has logged in successfully.
 | context: 
 | {
 |    "UserName": "Bob",
 |    "UserID": 35353,
 |    "Operation": "create",
 |    "environment": "production",
 |    "user_role": "admin"
 | }
 | extra: 
 | {
 |    "session-UID": "A09A248CF0",
 |    "memory": {
 |       "memory-used": "65%",
 |       "memory-total": "15,23",
 |       "memory-available": "5,30"
 |    },
 |    "CPU-used": "21,4%"
 | }
```

## Push pre-processors into logger

The factory provides `pushProcessors` methods used to add pre-processor into a given logger. 

```vbscript
Dim Logger As VBAMonologger.Logger
Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA()

' Add pre-processors UID
VBAMonologger.Factory.pushProcessorUID Logger, 8

' Add pre-processors CPU usage
VBAMonologger.Factory.pushProcessorUsageCPU Logger

' Add pre-processors Memory usage
VBAMonologger.Factory.pushProcessorUsageMemory Logger

' Add pre-processors Tags
Dim tags As Object
Set tags = CreateObject("Scripting.Dictionary")
tags.Add "environment", "production"
VBAMonologger.Factory.pushProcessorTags Logger, tags, TAGS_DESTINATION.LOG_EXTRA

' Use logger
Logger.trace "Authentication function call."
```

``` title='Result'
[2024/12/16 18:34:09] DEBUG: Authentication function call. | extra: {"session-UID":"F056D5EE","CPU-used":"0,0%","memory-used":"62%","tags":{"environment":"production"}}
```


## Show context and extra data on multiples lines (whitespace)

By default, the line formatter writes context and extra data on the same line, but you can configure it to display them on multiple lines. To achieve this, it is necessary to inject a custom formatter.

```vbscript
Public Sub howto_show_context_and_extra_on_multilples_line()
    ' Create a custom formatter
    Dim customFormatter As VBAMonologger.FormatterLine
    Set customFormatter = VBAMonologger.Factory.createFormatterLine
    customFormatter.showContext = True
    customFormatter.showExtra = True
    customFormatter.setTemplateLineWithNewlineForContextAndExtra
    customFormatter.withWhitespace = True
    customFormatter.withAllowingInlineLineBreaks = True
    
    ' Inject custom formatter
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA( _
        "App", _
        customFormatter _
    )

    ' Add pre-processors 
    VBAMonologger.Factory.pushProcessorUID Logger, 8
    VBAMonologger.Factory.pushProcessorUsageCPU Logger
    VBAMonologger.Factory.pushProcessorUsageMemory Logger
    Dim tags As Object
    Set tags = CreateObject("Scripting.Dictionary")
    tags.Add "environment", "production"
    VBAMonologger.Factory.pushProcessorTags _ 
        Logger, tags, TAGS_DESTINATION.LOG_EXTRA
    
    ' Set a dummy context
    Dim context As Object: Set context = CreateObject("Scripting.Dictionary")
    context.Add "Username", "v20100v"
    
    ' Use logger
    Logger.trace "Authentication function call."
    Logger.info "Adding the new user: '{username}'", context
End Sub
```

``` title='Result'
[2024/12/16 18:46:28] App.DEBUG: Authentication function call.
 | extra: 
 | {
 |    "session-UID": "F65A0049",
 |    "CPU-used": "0,0%",
 |    "memory-used": "61%",
 |    "tags": {
 |       "environment": "production"
 |    }
 | }
[2024/12/16 18:46:28] App.INFO: Adding the new user: '{username}'
 | context: 
 | {
 |    "Username": "v20100v"
 | }
 | extra: 
 | {
 |    "session-UID": "F65A0049",
 |    "CPU-used": "0,0%",
 |    "memory-used": "61%",
 |    "tags": {
 |       "environment": "production"
 |    }
 | }
```