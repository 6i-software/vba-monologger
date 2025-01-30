---
description: Learn how to inject loggers into VBA classes module, enhancing modularity and testability in your VBA applications.
---

## Dependency injection?

Dependency Injection is a design pattern used to implement IoC (Inversion of Control), allowing a class's dependencies to be injected into it rather than the class creating them itself. This software design technique enhances flexibility and modularity in applications. It involves providing the dependencies required by a class at instantiation time, rather than creating them within the class.

When a class implements the LoggerAware interface, it allows for a logger object to be injected into it from the outside, typically by a dependency injection container. This way, the class becomes flexible and testable, as you can easily swap out the logger for another implementation or a mock for testing purposes. 

## Logger aware interface

The `VBAMonlogger.Log.LoggerAwareInterface` is an interface used to enable dependency injection for logging capabilities in a custom class. Its purpose is to standardize how a logger is injected into a class. This allows any class that implements the interface to receive a logger instance, which it can then use to log messages. In this case, the dependency is the logger.

### Implementation of `LoggerAwareInterface`

Here's an example of a class implementing the `VBAMonologger.LoggerAwareInterface` interface:

```vbscript title="DummyClassUsingLoggerAware.cls"
Option Explicit

Implements VBAMonologger.LoggerAwareInterface

Private Logger As VBAMonologger.LoggerInterface

' -------------------------------------- '
'  Implementation: LoggerAwareInterface  '
' -------------------------------------- '
'@inheritdoc
Private Sub LoggerAwareInterface_setLogger(paramLogger As LoggerInterface)
    Set Logger = paramLogger
End Sub

'@Description("Proxy method for public exposition.")
Public Sub setLogger(paramLogger As LoggerInterface)
    Call LoggerAwareInterface_setLogger(paramLogger)
End Sub

' ---------------- '
'  Public methods  '
' ---------------- '
Public Sub foo()
    ' Minimal exemple of using the "injected" logger
    Logger.info "I believe I can fly!"
    
    ' Using the "injected" logger with a log context and placeholders
    Dim context As Object: Set context = CreateObject("Scripting.Dictionary")
    context.Add "User", "Bob"
    context.Add "Operation", "fly"
    Logger.notice "I believe {User} can {Operation} in the sky", context
End Sub
```

### Injecting a real logger into custom object

To inject a logger into an object, you can use the following code:

```vbscript title="DummyModule.bas"
Public Sub howto_use_dependency_injection_logger()
    Dim myDummyClass As New DummyClassUsingLoggerAware
    
    ' Inject a logger into object
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA("App")
    Call myDummyClass.setLogger(Logger)
    
    ' In foo method of DummyClass, we use the logger
    Debug.Print "=== Use the logger by dependencies injection ==="
    myDummyClass.foo
    Debug.Print ""
End Sub
```
``` title="Result"
=== Use the logger by dependencies injection ===
[2024/12/19 11:31:18] App.INFO: I believe I can fly!
[2024/12/19 11:31:19] App.NOTICE: I believe Bob can fly in the sky | context: {"User":"Bob","Operation":"fly"}
```