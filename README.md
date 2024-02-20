# pythonclrhost
## CLR Hosting in raw python

`pythonclrhost` enables CLR hosting of a dotnet runtime in python. It is currently only meant to wrap a [custom implementation of clr-loader](https://github.com/rkbennett/clr-loader)'s netfx binary. In the future I do plan on expanding this to support dynamic calls

## Basic Usage

### Host ClrLoader.dll in dotnet runtime and Get a function pointer from Python.Runtime.dll (pythonnet)
```python
from pythonclrhost import clrhost
clr = clrhost()
# Get ClrLoader.dll bytes (can be from a web request, etc.)
dll = open(r'c:\Path\To\ClrLoader.dll', 'rb').read()
# Start the runtime and load in our dll
clr.load(dll, "v4.0.30319")
# Initialize ClrLoader
clr.pyclr_initialize()
# Get Python.Runtime.dll bytes (from pythonnet, can also be from any source that returns bytes)
asm = open(r'c:\Path\To\Python.Runtime.dll', 'rb').read()
# Get a valid intpointer for the Initialize function from Python.Runtime.Loader
intPtr = clr.get_function(asm)
```

### Close all domain in ClrLoader
```python
from pythonclrhost import clrhost
clr = clrhost()
# Get ClrLoader.dll bytes (can be from a web request, etc.)
dll = open(r'c:\Path\To\ClrLoader.dll', 'rb').read()
# Start the runtime and load in our dll
clr.load(dll, "v4.0.30319")
# Initialize ClrLoader
clr.pyclr_initialize()
# Get Python.Runtime.dll bytes (from pythonnet, can also be from any source that returns bytes)
asm = open(r'c:\Path\To\Python.Runtime.dll', 'rb').read()
# Get a valid intpointer for the Initialize function from Python.Runtime.Loader
intPtr = clr.get_function(asm)
# Close all appdomains in ClrLoader (do not do this unless you are done interacting with this instance)
clr.pyclr_finalize()
# Remove the clr instance just to be safe
del clr
```

### Create an app domain in ClrLoader
```python
from pythonclrhost import clrhost
clr = clrhost()
# Get ClrLoader.dll bytes (can be from a web request, etc.)
dll = open(r'c:\Path\To\ClrLoader.dll', 'rb').read()
# Start the runtime and load in our dll
clr.load(dll, "v4.0.30319")
# Initialize ClrLoader
clr.pyclr_initialize()
# Create AppDomain and return ClrLoader index reference (This can be passed into Get function later if you load an assembly into it)
domain = clr.create_appdomain("FooBarBaz")
```

### Closing an app domain in ClrLoader
```python
from pythonclrhost import clrhost
clr = clrhost()
# Get ClrLoader.dll bytes (can be from a web request, etc.)
dll = open(r'c:\Path\To\ClrLoader.dll', 'rb').read()
# Start the runtime and load in our dll
clr.load(dll, "v4.0.30319")
# Initialize ClrLoader
clr.pyclr_initialize()
# Create AppDomain and return ClrLoader index reference (This can be passed into Get function later if you load an assembly into it)
domain = clr.create_appdomain("FooBarBaz")
# Close the created AppDomain
clr.close_appdomain(domain)
```

### Gotchas

This is still a work in progress, as such there is still development being done for making autoresolution of valid runtimes work via `EnumerateInstalledRuntimes`. There is also still work being done to allow for dynamic function arguments. This requires a decent amount of work because all arguments have to be put into a SafeArray, which has to be explicitly defined before execution.