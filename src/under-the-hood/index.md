---
hide:
  - footer
---

# Under the Hood

The idea is to dive deep, step by step, and see what’s happening under the hood of **VBA Monologger**. Each component plays a vital role in ensuring logs are structured, meaningful, handled and routed effectively.  

Below are the key components and their roles:

<div class="grid cards" markdown>

- :material-archive-outline: __[Log levels]__  
  Classify the importance of log messages, ranging from detailed debugging information to critical errors that require immediate attention.

- :material-content-save: __[Log record]__  
  Captures detailed information about an event or action within an application, forming the core of a log entry.

- :material-format-paint: __[Formatter]__  
  Structures log messages into a specific format, making them readable and suited for various outputs such as JSON, text, or HTML.

- :material-console-network-outline: __[Handler]__  
  Routes log records to their appropriate destinations, such as files, console outputs, or external systems.

- :material-shape-plus: __[Pre-Processor]__  
  Adds or modifies context in log entries before they are formatted or handled, providing flexibility and enrichment to the log data.

- :material-engine-outline: __[Logger]__  
  The central component that orchestrates logging by managing levels, handlers, formatters, and pre-processors to ensure seamless log management.

</div>

[Log levels]: log-severity-levels.md
[Log record]: log-record.md
[Formatter]: formatter.md
[Handler]: handler.md
[Pre-Processor]: pre-processor.md
[Logger]: logger.md

