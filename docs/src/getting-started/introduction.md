## Preamble

### A logging system

Logging is a process that involves recording and storing traces of events, activities, or errors that occur during the use of an application. Useful for both developers for debugging and administrators for diagnosing and resolving incidents, logs provide traceability and visibility into the behavior of an application throughout its operation.

In its simplest form, log entries are recorded in text format, with each line representing an event that occurred during the application's lifecycle.

```bash
[2024-11-05 09:15:34] app.INFO: Application started
  | context: {"version": "2.3.1", "user_id": 101}
[2024-11-05 09:16:01] app.INFO: Workbook loaded successfully
  | context: {"workbook_id": 789, "workbook_name": "Q4 Marketing"}
[2024-11-05 09:17:15] app.DEBUG: Task modified
  | context: {"task_id": 456, "task_name": "Strategy Review", "user_id": 101}
[2024-11-05 09:18:45] app.WARNING: Low disk space - risky save
  | context: {"available": "100 MB", "required": "200 MB"}
[2024-11-05 09:19:03] app.ERROR: Project save failed
  | context: {"project_id": 789, "error": "Insufficient disk space", "user_id": 101}
[2024-11-05 09:20:00] app.INFO: Application closed
  | context: {"user_id": 101}
```

According to the *twelve-factor app* manifest, you should "*[treat logs as event streams](https://12factor.net/logs)*". Logs are not just recorded in a file for later consultation. They can also be monitored in real-time in a terminal, sent to a database, or redirected to external log aggregation and analysis tools (such as the ELK stack, Graylog, BetterStack, Splunk...).

A logging system should offer flexible management, allowing different severity levels of events to be distinguished, so messages can be filtered according to their importance, from simple information to critical errors. It should also be capable of sending logs to multiple destinations simultaneously, such as a terminal, file, database, or monitoring service. Additionally, the log format must be customizable to meet the specific needs of the application and the tools used for analysis, making it easier to manage and interpret the collected data.


### Motivations 

VBA provides developers with the ability to automate tasks, interact with the features of Microsoft Office applications, and even create applications with a graphical user interface (`Userform`). However, compared to other development ecosystems, VBA only offers a rudimentary logging solution, limited to the `Debug.Print` function, which writes to the Excel console (a.k.a. the Excel immediate window).

The *VBA Monologger* library project was born out of the need for a more advanced and flexible logging solution in the VBA ecosystem. It is (heavily) inspired by the PSR-3 standard in the PHP ecosystem and its most recognized implementation, the [Monolog](https://github.com/Seldaek/monolog) library, which itself is inspired by the [Logbook](https://logbook.readthedocs.io/en/stable/) library in Python.

The goal of this library is to provide similar features and capabilities, particularly by offering a modular architecture that can easily adapt to different use cases. 


### What is VBA Monologger?

VBA Monologger is an advanced and flexible logging solution for VBA (*Visual Basic for Applications*) ecosystem.

Main features:

- Customize the logging format to define how log messages are structured and displayed.
- Specify the destination where logs should be viewed (e.g., VBA console *a.k.a* Excel's immediate window, Windows console (cmd.exe) with ANSI color support, file...) and configure the conditions under which logging events are triggered based on specific criteria.
- Manages 8 standard severity levels to classify the importance of log messages, following the PSR-3 standard.
- Enrich log records with pre-processors, enabling the addition of context, transformation of data, or customization of log entries to suit specific needs (e.g. add CPU or memory usage, generate a UID for each session, add tags... and more).
- Use the provided loggers in the library (*e.g.* `LoggerConsoleVBA`, `LoggerConsole` or `LoggerFile`) for basic usage, or create your own custom logging system.
- Easily develop your own custom formatter, handler, and pre-processors to tailor the logging system to your specific needs. By creating unique formatting styles, specialized handlers, and custom pre-processing logic, you can enhance the functionality and flexibility of your logging setup, ensuring it meets the precise requirements of your applications and workflows.


## Concepts

### Severity log level

It indicates the severity of each event, from the most trivial to the most catastrophic, and allow administrators or developers to filter messages based on their importance.

*VBA Monologger* manages 8 standard severity levels **to classify the importance of log messages**, following the [PSR-3](https://www.php-fig.org/psr/psr-3/) standard, which is itself based on [RFC-5424](https://www.rfc-editor.org/rfc/rfc5424), the standard defined by the IETF (*Internet Engineering Task Force*) to specify the format of messages for the Syslog protocol, which is used for transmitting logs over IP networks.

| Log level   | Description                                                                                                                                                                                                                           |
|-------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `EMERGENCY` | Indicates a very critical situation that requires immediate attention. (*system crash, data corruption*)                                                                                                                              |
| `ALERT`     | Signals an alert condition. (*critical disk space running out*)                                                                                                                                                                       |
| `CRITICAL`  | Indicates a serious error. (*database connection failure, server downtime*)                                                                                                                                                           |
| `ERROR`     | Represents an error in the system. (*failed to save user data, unexpected exception*)                                                                                                                                                 |
| `WARNING`   | A warning about a potential problem. (*use a deprecated function used, low memory warning*)                                                                                                                                           |
| `NOTICE`    | Important notifications that are not urgent. (*user login successful, configuration change detected*)                                                                                                                                 |
| `INFO`      | General information about the normal operation. (*system startup, data processed successfully*)                                                                                                                                       |
| `TRACE`     | Detailed information for debugging. (*variable values during loop iteration, execution details*). Notes, that the '**debug**' method exposes presents in PSR-3 is rename into '**trace**' in order to be compatible in VBA ecosystem. |



### Logger

**The logger is the central component of this library, acting as the primary interface for recording, categorizing, and managing log messages throughout an application**. It provides developers with a highly configurable and flexible tool for implementing custom logging logic tailored to their specific needs. By using a logger, applications can systematically capture events and system states, facilitating both real-time monitoring and historical analysis of system behavior.

The logger is designed to handle multiple logging levels, directing each log entry to the appropriate handlers (*i.e.* appropriate destinations) and applying the correct formatting to messages. It also supports the use of various pre-processors, which can enrich log messages with extra contextual information, allowing for complex logging logic while keeping code readable and straightforward.

```mermaid
mindmap
  root((Logger))
    node1(Handlers)
      HandlerConsoleVBA<br>*for all levels*
        FormatterLine
      HandlerConsole<br>*for all levels*
        FormatterANSIColoredLine
      HandlerFile<br>*exclude debug level*
        FormatterJSON
      HandlerEmail<br>*for level greater than error*
        FormatterHTML
    node2(Processors)
      ProcessorPlaceholders
      ProcessorUID
    node3(Name of loger, a.k.a. log channel)
```

Additionally, the logger standardizes and simplifies the use of logging methods (such as methods: `logger.trace`, `logger.info`, ...). It offers a consistent and intuitive approach to logging at different levels of severity, letting developers effortlessly call the appropriate logging level without dealing with the underlying technical details. Each log level can be invoked through a simple, clear method, making logging an integral yet unobtrusive part of the development process. 

Every logger implements the `LoggerInterface`, which provides the following methods:

```vbscript
Logger.emergency "A critical failure occurred in the application."
Logger.alert "Action required: unable to generate the dashboard."
Logger.critical "System is in an unstable state."
Logger.error "An error occurred with the user 'DRZCFOS2'."
Logger.warning "'Beetlejuice' should not be called more than 3 times."
Logger.notice "Process completed successfully with minor issues."
Logger.info "User 'UltraVomit' has logged in successfully."
Logger.trace "Authentication function call for user 'Bob Morane'." 
```


### Identifying a logger with a channel

**A channel is a powerful way to identify which part of an application a log entry is associated with**. This is especially useful in large applications with multiple components and multiple loggers. The idea is to have several logging systems sharing the same handler, all writing into a single log file. Channels help identify the source of the log, making filtering and searching more manageable.

Here’s an example with three distinct logging channels to demonstrate how they help differentiate logs by application component: one channel for the main application (`app`), another for authentication (`auth`), and a third for data processing (`data`).

``` title='Logging system with multiples loggers identified by channels'
[2024-11-05 09:15:34] auth.INFO: User login successful
[2024-11-05 09:16:01] app.INFO: Dashboard loaded successfully
[2024-11-05 09:16:20] data.DEBUG: Data import started
[2024-11-05 09:17:30] auth.WARNING: Suspicious login attempt detected
[2024-11-05 09:18:45] data.ERROR: Data import failed
[2024-11-05 09:19:03] app.INFO: User preferences saved
[2024-11-05 09:20:00] app.INFO: Application shutdown initiated
```


### Processing log records with a handler

**A log handler is a key component responsible for processing each log entry**. When a log message is generated, it is not simply recorded; it must be directed to a location where it can be viewed and used. This is where the handler comes in, determining where and how each log entry will be sent or saved. 

Here are some examples of *built-in* log handlers provided in VBAMonologger:

| Handler             | Description                                                                                                             |
|---------------------|-------------------------------------------------------------------------------------------------------------------------|
| `HandlerConsoleVBA` | Sends log messages to the console of VBA Project IDE (*Excel's Immediate Window*).                                      |
| `HandlerConsole`    | Sends log messages to the Windows console (*cmd.exe*).                                                                  |
| `HandlerFile`       | Write log messages into a text file.                                                                                    |
| ~~`HandlerEmail`~~    | ~~Sends messages by email, typically used to alert an administrator in case of critical errors.~~ (*not yet available*) |

The benefit of using different handlers lies not only in applying specific treatments to logs but also in filtering messages based on their severity level. A handler can be configured to handle only certain severity levels. For example, one handler could be set to log only critical errors to a dedicated file, while another handler records all events in a general log file.


### Formatting log records, the serialization of logs record

**Each *handler* is associated with a unique *formatter*, a specialized component responsible for defining the structure and presentation of log messages**. The formatter processes and organizes each log entry, converting it from its raw form to a readable format tailored to a specific type (text, HTML, JSON, etc.). This process can be seen as similar to the serialization of a log record.

*VBA Monologger* provides the following formatters:

| Log Formatter            | Description                                                                                                                                                     |
|--------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `FormatterLine`          | The default formatter that represents each log entry on a single line of text.                                                                                  |
| `FormatterANSIColoredLine` | A version derived from *FormatterLine* that supports color coding each log entry line using ANSI escape sequences.                                              |
| `FormatterJSON`          | Formats the logs in JSON. This is the most interoperable format, facilitating integration with external log aggregation and analysis tools (e.g., *ELK stack*). |
| ~~`FormatterHTML`~~        | ~~Produces messages in HTML format, typically used when sending logs via email.~~ (*not yet available*)                                                         |


### Redirection and chaining handlers (*stack handlers*)

**In the log processing, there's no limitation to having multiple handlers into logger**, so that the same log entry can be sent to multiple destinations in the same time with different formatter: console, writing to a file, or sending via email. Each handler acts sequentially, one after the other, in the order they were added to the logger's stack. 

When a log event occurs, it passes through all handlers, each performing its own processing.

This mechanism provides great flexibility in log processing because each handler can be independently configured to perform specific actions without interfering with the others.


### Propagation control of logs (*bubbling*)

**When multiple handlers are referenced in a logger, each handler can choose either to block the propagation of a log message (*a.k.a.* bubbling) within the processing chain or to allow it to continue to other handlers in the stack.** The control of this propagation is managed by setting the `bubble` boolean property attached to each *handler*. 

When a *handler* blocks propagation (*i.e.* `bubble = false`), it means the log message will not be passed to the *handlers* below it in the stack. Otherwise, the message will continue to propagate until every *handler* in the stack has had a chance to process it.

Let's imagine a logging system with three *handlers*:

- First, a `HandlerEmail` to send error-level logs (`level >= ERROR`) via email, without propagation (`bubble = false`).
- Next, a `HandlerFile` to log other messages in a file, excluding the trace level (`TRACE < level < ERROR`).
- And finally, a `HandlerConsole` to display the remaining logs in the console (`level < ERROR`).

With this custom logging system, when an error-level log is captured by `HandlerEmail`, the first handler in the stack, it will not propagate to the other *handlers*. An `ERROR` level log will only be sent via email. It will not be recorded in the log file (`HandlerFile`), nor displayed in the console (`HandlerConsole`).

```mermaid
sequenceDiagram
    participant Logger as Logger
    participant HandlerEmail as HandlerEmail<br>($bubble = false, level = ERROR)
    participant HandlerFile as HandlerFile<br>(level = NOTICE)
    participant HandlerConsole as HandlerConsole<br>(level = DEBUG)

    %% ERROR level log
    Logger->>HandlerEmail: ERROR level log
    HandlerEmail-->>HandlerEmail: Send log via email
    HandlerEmail--xHandlerFile: Propagation blocked<br>($bubble = false)
    Note right of HandlerEmail: The error message is handled <br>only by HandlerEmail <br>and is not processed elsewhere.
```

And when a log of level `INFO` is recorded, it is not captured by the first `HandlerEmail`. Its processing starts with the `HandlerFile`, which allows the propagation of the messages. The log is then sent to the `HandlerConsole` for processing. Therefore, this log will be recorded in a log file and displayed in the console.

```mermaid
sequenceDiagram
    participant Logger as Logger
    participant HandlerEmail as HandlerEmail<br>($bubble = false, level = ERROR)
    participant HandlerFile as HandlerFile<br>(level = NOTICE)
    participant HandlerConsole as HandlerConsole<br>(level = DEBUG)

    %% INFO level log
    Logger->>HandlerEmail: INFO level log
    HandlerEmail--xHandlerEmail: Ignored (insufficient level)
    HandlerEmail->>HandlerFile: INFO level log
    HandlerFile-->>HandlerFile: Log saved to file
    HandlerFile-->>HandlerConsole: Propagation (bubble = true)
    HandlerConsole-->>HandlerConsole: Displayed in console
    Note right of HandlerFile: INFO level log is processed<br>by HandlerFile and HandlerConsole.
```


### Adding metadatas in log records

In addition to the basic log message, you may sometimes want to include extra information that helps to provide more context for the event being logged. This could include things like the username of the person triggering the event, a session ID, or any other piece of data that can assist in understanding the log entry better. 

*VBA Monologger* offers two options for adding complementary data to a log record: `context` and `extra`.

 - The `context` is used to add information directly related to the logged event, such as details about an error or an ongoing operation.
 - Meanwhile, `extra` is reserved for additional metadata, often generated automatically or added by pre-processors, providing a global context.

Whether it's the `context` option or the `extra` option, they are essentially VBA dictionaries, where you can store key-value pairs that hold relevant information. When you create a log entry, this context can be attached and will be incorporated into the log output, providing deeper insights into the logged event. This feature is a great way to enrich your log entries with important details and provide better traceability and understanding of your application's behavior.

```vbscript title='Using placeholders with data context'
' Set context 
Dim context As Object: Set context = CreateObject("Scripting.Dictionary")
context.Add "Username", "v20100v"

' Set extra 
Dim extra As Object: Set extra = CreateObject("Scripting.Dictionary")
extra.Add "CPU-Usage", "51%"

Logger.info "Adding a new user", context
Logger.info "Adding the new user: '{username}'", context
Logger.info "Adding the new user: '{username}'", context, extra
```

``` title='Result'
[2024-11-05 09:15:34] app.INFO: Adding a new user | {"Username": "v20100v"}
[2024-11-05 09:15:34] app.INFO: Adding the new user: 'v20100v' | {"Username": "v20100v"}
[2024-11-05 09:15:34] app.INFO: Adding the new user: 'v20100v' | {"Username": "v20100v"} | extra: {"CPU-Usage":"51%"}
```


### Template engine with placeholders

Regardless of which log handler is used and which formatter is applied, the fields within the `context` can be accessed as template variables within the log message. It is not possible to use the data present in `extra`; only the data in `context` is utilized by the template engine.

This templating engine is provided by the pre-processor `ProcessorPlaceholders`, which uses `context` data to replace variables in log messages. It allows to embed context values directly into log message text with placeholders. All template variables are enclosed in `{}` brackets. 

Notes, if a particular key doesn’t exist in the context, it will be replaced by an empty string.

In the example above, the engine automatically replaces the placeholder `{username}` with its corresponding value from the context dictionary, in this case, `v20100v`.



### Pre-processor of log records

Pre-processors are a powerful feature, allowing for additional metadatas to be added to log messages before they are recorded. These functions can be used to enrich log messages with extra information that might not be directly part of the log entry itself, but is still relevant for better understanding and tracking. Pre-processors can modify, format, or even generate additional metadata that can be attached to each log message into the `extra` property. A logger can reference one or more pre-processors.

In VBA Monologger, several built-in processors offer specific functionalities to enhance the log entry by adding additional context or modifying it in various ways before it is passed to the handlers. Here are some examples of the available pre-processors:

| Log processor           | Description                                                                                                                                                                                                                                                                                               |
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