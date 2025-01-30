---
description: Learn how to create a null logger in VBA Monologger.
---

## Why choose a null logger?

Using a null Logger can be quite useful in several situations:

1. **In testing or development**: When writing tests or developing applications, you might not need logs. A Null Logger allows you to disable logging without changing your production code.
2. **Selective log disabling**: Sometimes, you may want to disable logging for specific parts of your application without completely removing log calls. The null Logger can be injected where logging is unnecessary.
3. **Performance**: In performance-sensitive environments, avoiding logging operations can reduce latency and improve performance.

In summary, the null Logger is a practical solution for managing logs flexibly and efficiently, allowing you to temporarily or conditionally disable logging without altering the rest of your code.


## Create a null logger

The factory provides a method to create a null logger.

```vbscript 
Public Sub howto_create_null_logger()
    ' Create a null logger instance that does nothing!
    Dim LoggerNull As VBAMonologger.LoggerInterface
    Set LoggerNull = VBAMonologger.Factory.createLoggerNull()
    
    ' Logs message for each severity levels
    LoggerNull.trace "Authentication function call for user 'Bob Morane'." 
    LoggerNull.info "User 'UltraVomit' has logged in successfully."
    LoggerNull.notice "Process completed successfully with minor issues."
    LoggerNull.warning "'Beetlejuice' should not be called more than 3 times."
    LoggerNull.error "An error occurred with the user 'DRZCFOS2'."
    LoggerNull.critical "System is in an unstable state."
    LoggerNull.alert "Action required: unable to generate the dashboard."
    LoggerNull.emergency "A critical failure occurred in the application."
End Sub
```