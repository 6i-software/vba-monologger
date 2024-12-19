## What is a log context?

In addition to the basic log message, you may sometimes want to include [extra data](../introduction.html#adding-metadatas-in-log-records) that helps to provide more context for the event being logged. You can give a variable context with log message. It is a simply VBA dictionary, where you can store key-value pairs that hold relevant information. When you create a log entry, this context can be attached and will be incorporated into the log output, providing deeper insights into the logged event.

This variable can simply be displayed (or not) into the log message, or can be consumed by using placeholders in log message.


## Create a log context

To create a log context (*i.e.* a VBA dictionary), it is recommended to do it like this:

```vbscript
' Set context 
Dim context As Object: Set context = CreateObject("Scripting.Dictionary")
context.Add "Username", "v20100v"
```


## Use placeholders into log message

The fields within the context can be accessed as template variables within the log message. This templating engine is provided by the pre-processor `ProcessorPlaceholders`, which uses context data to replace variables in log messages. It allows to embed context values directly into log message text with placeholders. All template variables are enclosed in `{}` brackets.

Notes, if a particular key doesn’t exist in the context, it will be replaced by an empty string.

In the example, the engine automatically replaces the placeholder `{username}` with its corresponding value from the context dictionary, in this case, `v20100v`.

```vbscript
' Set context 
Dim context As Object: Set context = CreateObject("Scripting.Dictionary")
context.Add "Username", "v20100v"

' Only display log context (if configuration's formatter allow to show it)
Logger.info "Adding a new user", context

' Consume log context with placeholder
Logger.info "Adding the new user: '{username}'", context
```

``` title='Result'
[2024-11-05 09:15:34] app.INFO: Adding a new user | {"Username": "v20100v"}
[2024-11-05 09:15:34] app.INFO: Adding the new user: 'v20100v' | {"Username": "v20100v"}
```