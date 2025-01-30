---
description: Understand the central role of the logger in VBA Monologger.
---

## Concept

**The logger is the central component of this library, acting as the primary interface for recording, categorizing, and managing log messages throughout an application**. It provides developers with a highly configurable and flexible tool for implementing custom logging logic tailored to their specific needs. By using a logger, applications can systematically capture events and system states, facilitating both real-time monitoring and historical analysis of system behavior.

The logger is designed to handle multiple logging levels, directing each log entry to the appropriate handlers (*i.e.* appropriate destinations) and applying the correct formatting to messages. It also supports the use of various pre-processors, which can enrich log messages with extra contextual information, allowing for complex logging logic while keeping code readable and straightforward.

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

