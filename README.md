# VBA Monologger

[![LICENSE](https://img.shields.io/badge/license-MIT-informational.svg)](https://github.com/v20100v/6i-Jekyll/blob/develop/LICENSE.md)
&nbsp;[![Want to support me? Offer me a coffee!](https://img.shields.io/badge/Want%20to%20support%20me%3F%20Offer%20me%20a%20coffee%21-donate-informational.svg)](https://www.buymeacoffee.com/vincent.blain)

> VBA Monologger is an advanced and flexible logging solution for VBA (*Visual Basic for Applications*) ecosystem. Easily send logs to the Excel console (Immediate Windows), or simultaneously to a file, or even in a Windows console. Set up in two minutes. It is largely inspired by the [Monolog](https://github.com/Seldaek/monolog) library in PHP, which itself is inspired by the [Logbook](https://logbook.readthedocs.io/en/stable/) library in Python.
>
> 📕 Documentation : [https://6i-software.github.io/vba-monologger/](https://6i-software.github.io/vba-monologger/)
 

## Introduction

VBA provides developers with the ability to automate tasks, interact with the features of Microsoft Office applications, and even create applications with a graphical user interface (`Userform`). However, compared to other development ecosystems, VBA only offers a rudimentary logging solution, limited to the `Debug.Print` function, which writes to the Excel console (a.k.a. the Excel immediate window).

The *VBA Monologger* library project was born out of the need for a more advanced and flexible logging solution in the VBA ecosystem. It is (heavily) inspired by the PSR-3 standard in the PHP ecosystem and its most recognized implementation, the Monolog library. The goal of this library is to provide similar features and capabilities, particularly by offering a modular architecture that can easily adapt to different use cases. The main idea is for each developer to easily configure and customize their own logging system according to their needs.


## Features

**In VBA Monologger, the logger is the central component of this library, acting as the primary interface for recording, categorizing, and managing log messages throughout an application**. It provides developers with a highly configurable and flexible tool for implementing custom logging logic tailored to their specific needs. By using a logger, applications can systematically capture events and system states, facilitating both real-time monitoring and historical analysis of system behavior.

The logger is designed to handle multiple logging levels, directing each log entry to the appropriate handlers (i.e. appropriate destinations) and applying the correct formatting to messages. It also supports the use of various pre-processors, which can enrich log messages with extra contextual information, allowing for complex logging logic while keeping code readable and straightforward.

Main features:

- Customize the logging format to define how log messages are structured and displayed.
- Specify the destination where logs should be viewed (*e.g.*, VBA console *a.k.a* Excel's immediate window, Windows console (cmd.exe) with ANSI color support, file...) and configure the conditions under which logging events are triggered based on specific criteria.
- Manages 8 standard severity levels to classify the importance of log messages, following the [PSR-3](https://www.php-fig.org/psr/psr-3/) standard.
- Enrich log records with pre-processors, enabling the addition of context, transformation of data, or customization of log entries to suit specific needs (*e.g.* add CPU or memory usage, generate a UID for each session, add tags... and more).
- Use the provided loggers in the VBAMonologger factory (*e.g. `LoggerConsoleVBA`, `LoggerConsole` or `LoggerFile`*) for basic usage, or create your own custom logging system.
- Easily develop your own custom formatter, handler, and pre-processors to tailor the logging system to your specific needs. By creating unique formatting styles, specialized handlers, and custom pre-processing logic, you can enhance the functionality and flexibility of your logging setup, ensuring it meets the precise requirements of your applications and workflows.


## Documentation

Please refer to the documentation for details on how to install and use VBA Monologger.
> [Website documentation](https://6i-software.github.io/vba-monologger/)


## Quick start

### Manual installation

1. Download the VBA Monologger Excel Add-in (.xlam file) to your computer:   [6i_VBA-Monologger.xlam](https://github.com/6i-software/vba-monologger/raw/refs/heads/main/src/6i_VBA-Monologger.xlam)
2. Put this xlam file into a folder trusted by Excel, and add it as a reference in your VBA project through *Tools > References* in the VBA editor.

### Log output to VBA Console

In VBA Monologger, we use a factory in order to simplify and standardize the creation of objects, such as loggers, by encapsulating the logic needed to initialize them. The factory pattern abstracts the object creation process, which can be particularly useful. So to instantiate your first logger that output logs into the VBA console, just use the method `VBAMonologger.Factory.createLoggerConsoleVBA()`, as shown below. It provides an instance of a logger preconfigured with default handler, formatter, and pre-processors.

```vbscript
Public Sub howto_use_loggerConsoleVBA()
    ' Create a logger instance for output log into VBA console (Excel's immediate window)
    Dim Logger As VBAMonologger.LoggerInterface
    Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA("App")
    
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

You can see result in the VBA console (a.k.a. Excel's Immediate Windows).

![VBAMonologger-output-VBAConsole.png](https://6i-software.github.io/vba-monologger/getting-started/VBAMonologger-output-VBAConsole.png)

As you can see, in the signature of this factory's method, it is possible to set the name the logger (channel) and to load a custom formatter.

```vbscript
Public Function createLoggerConsoleVBA( _
    Optional ByVal paramLoggerName As String = vbNullString, _
    Optional ByRef paramFormatter As VBAMonologger.FormatterInterface = Nothing _
) As VBAMonologger.Logger
``` 


### Log output to Windows console

If you prefer to display your logs outside the Excel VBA IDE, you can output them directly to the Windows Console (cmd.exe). 

The factory can create a dedicated logger for Windows Console with `VBAMonologger.Factory.createLoggerConsole()` method. It handles log messages by streaming them to the Windows console using an HTTP-based client/server architecture. The client sends log records as HTTP requests to the server, and the server processes these requests, displaying the log messages directly in the console output. This logger features a formatter that supports ANSI colors `VBAMonologger.Formatter.FormatterANSIcoloredLine`. 

It also includes the pre-processors placeholders according to PSR-3 rules. It allows to use placeholders in log message that will be replaced with value provided in the log record context.


```vbscript
Public Sub howto_use_logger_console()
    Dim Logger As VBAMonologger.LoggerInterface
    Set Logger = VBAMonologger.Factory.createLoggerConsole("App")

    ' Use the logger for each severity levels
    Logger.trace "Authentication function call for user 'Bob Morane'." 
    (...)
End Sub    
```

When you execute this code, it launches a `cmd.exe`, and you can view the results in it. The formatter's configuration allows you to customize the color scheme.

![VBAMonologger-output-WindowsConsole.png](https://6i-software.github.io/vba-monologger/getting-started/VBAMonologger-output-WindowsConsole.png)


### Log output to file

You can send logs into a file with the default logger file provided by factory's method `VBAMonologger.Factory.createLoggerFile()`. By default, this logger writes logs to the `./var/log/logfile_yyyy-mm-dd.log` file, relative to the path of the workbook. To ensure compatibility with special and multilingual characters, the UTF-8 encoding is preferred.

```vbscript
Public Function createLoggerFile( _
    Optional ByVal paramLoggerName As String = vbNullString, _
    Optional ByRef paramFormatter As FormatterInterface = Nothing, _
    Optional ByVal paramLogFileName As String = vbNullString, _
    Optional ByVal paramLogFileFolder As String = vbNullString _
) As VBAMonologger.Logger
```

Here’s an example with a custom name and custom folder of the log file.

```vbscript
Public Sub howto_change_logger_file_name_and_folder()
    Dim Logger As VBAMonologger.LoggerInterface
    Set Logger = VBAMonologger.Factory.createLoggerFile( _ 
        paramLoggerName:= "App", _
        paramLogFileName:="my-log-file___"&format(Now, "yyyy-mm-dd") & ".log", _
        paramLogFileFolder:=ThisWorkbook.Path & "\logs" _        
    )
    
    ' Logs message for each severity levels
    Logger.trace "Authentication function call for user 'Bob Morane'."
    (...)
End Sub
```

![VBAMonologger-output-File.png](https://6i-software.github.io/vba-monologger/getting-started/VBAMonologger-output-File.png)


## About

### Acknowledgements

This library is (heavily) inspired by the work of [Jordi Boggiano, a.k.a. Seldaek](https://github.com/Seldaek), with [Monolog](https://github.com/Seldaek/monolog). Knowing that Monolog is itself inspired by Python's [Logbook](https://logbook.readthedocs.io/en/stable/) library. Let's thank Seldaek, Armin Ronacher, Georg Brandl, and others for their work!

### Want to support me? Offer me a coffee!

**VBA Monologger** is free and open source under the [MIT License](./LICENSE), but if you want to support me, you can [offer me a coffee here](https://www.buymeacoffee.com/vincent.blain) or by scanning this QR code. Thank you in advance for your assistance (and your appreciation) for this work ^^.

<a href="https://www.buymeacoffee.com/vincent.blain"><img alt="Buy me a coffee ?" src="https://6i-software.github.io/vba-monologger/assets/v20100v_buy-me-a-coffee_qrcode.png" width="300" height="300" /></a>


### Want to contribute ?

Ideas, bug reports, reports a typo in documentation, comments, pull-request & Github stars are always welcome !


### License

Release under [MIT License](./LICENSE),<br/>
Copyright (c) 2024 by 2o1oo vb20100bv@gmail.com
