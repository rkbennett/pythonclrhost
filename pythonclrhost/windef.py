import ctypes

# BindingFlag ints
class BindingFlags:
     BindingFlags_Default = 0
     BindingFlags_IgnoreCase = 1
     BindingFlags_DeclaredOnly = 2
     BindingFlags_Instance = 4
     BindingFlags_Static = 8
     BindingFlags_Public = 16
     BindingFlags_NonPublic = 32
     BindingFlags_FlattenHierarchy = 64
     BindingFlags_InvokeMethod = 256
     BindingFlags_CreateInstance = 512
     BindingFlags_GetField = 1024
     BindingFlags_SetField = 2048
     BindingFlags_GetProperty = 4096
     BindingFlags_SetProperty = 8192
     BindingFlags_PutDispProperty = 16384
     BindingFlags_PutRefDispProperty = 32768
     BindingFlags_ExactBinding = 65536
     BindingFlags_SuppressChangeType = 131072
     BindingFlags_OptionalParamBinding = 262144
     BindingFlags_IgnoreReturn = 16777216

# VARIANT TYPE flags
VT_EMPTY = 0
VT_NULL = 1
VT_I2 = 2
VT_I4 = 3
VT_R4 = 4
VT_R8 = 5
VT_CY = 6
VT_DATE = 7
VT_BSTR = 8
VT_DISPATCH = 9
VT_ERROR = 10
VT_BOOL = 11
VT_VARIANT = 12
VT_UNKNOWN = 13
VT_DECIMAL = 14
VT_I1 = 16
VT_UI1 = 17
VT_UI2 = 18
VT_UI4 = 19
VT_I8 = 20 # not allowed in VARIANT
VT_UI8 = 21 # not allowed in VARIANT
VT_INT = 22
VT_UINT = 23
VT_VOID = 24 # not allowed in VARIANT
VT_HRESULT = 25 # not allowed in VARIANT
VT_PTR = 26 # not allowed in VARIANT
VT_SAFEARRAY = 27 # not allowed in VARIANT
VT_CARRAY = 28 # not allowed in VARIANT
VT_USERDEFINED = 29 # not allowed in VARIANT
VT_LPSTR = 30 # not allowed in VARIANT
VT_LPWSTR = 31 # not allowed in VARIANT
VT_RECORD = 36
VT_FILETIME = 64 # not allowed in VARIANT
VT_BLOB = 65 # not allowed in VARIANT
VT_STREAM = 66 # not allowed in VARIANT
VT_STORAGE = 67 # not allowed in VARIANT
VT_STREAMED_OBJECT = 68 # not allowed in VARIANT
VT_STORED_OBJECT = 69 # not allowed in VARIANT
VT_BLOB_OBJECT = 70 # not allowed in VARIANT
VT_CF = 71 # not allowed in VARIANT
VT_CLSID = 72 # not allowed in VARIANT
VT_BSTR_BLOB = 0xfff # not allowed in VARIANT
VT_VECTOR = 0x1000 # not allowed in VARIANT
VT_ARRAY = 0x2000
VT_BYREF = 0x4000
VT_TYPEMASK = 0xfff

win_errors = {
    "2147500035": "E_POINTER"
}

# COM Interface Structures
class BSTR(ctypes._SimpleCData):
    _type_ = "X"

class SafeArrayBound(ctypes.Structure):
    _fields_ = [
        ("cElements", ctypes.c_uint32),
        ("lLbound", ctypes.c_int32)
    ]

class SafeArray(ctypes.Structure):
    _fields_ = [
        ("cDims", ctypes.c_uint16),
        ("fFeatures", ctypes.c_uint16),
        ("cbElements", ctypes.c_uint32),
        ("cLocks", ctypes.c_uint32),
        ("pvData", ctypes.c_void_p),
        ("rgsabound", SafeArrayBound * 1)
    ]

class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8)
    ]

class AssemblyVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("GetTypeInfoCount", ctypes.c_void_p),
        ("GetTypeInfo", ctypes.c_void_p),
        ("GetIDsOfNames", ctypes.c_void_p),
        ("Invoke", ctypes.c_void_p),
        ("get_ToString", ctypes.c_void_p),
        ("Equals", ctypes.c_void_p),
        ("GetHashCode", ctypes.c_void_p),
        ("GetType", ctypes.c_void_p),
        ("get_CodeBase", ctypes.c_void_p),
        ("get_EscapedCodeBase", ctypes.c_void_p),
        ("GetName", ctypes.c_void_p),
        ("GetName_2", ctypes.c_void_p),
        ("get_FullName", ctypes.c_void_p),
        ("get_EntryPoint", ctypes.c_void_p),
        ("GetType_2", ctypes.WINFUNCTYPE(
                                    ctypes.c_ulong, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p
        )),
        ("GetType_3", ctypes.c_void_p),
        ("GetExportedTypes", ctypes.c_void_p),
        ("GetTypes", ctypes.c_void_p),
        ("GetManifestResourceStream", ctypes.c_void_p),
        ("GetManifestResourceStream_2", ctypes.c_void_p),
        ("GetFile", ctypes.c_void_p),
        ("GetFiles", ctypes.c_void_p),
        ("GetFiles_2", ctypes.c_void_p),
        ("GetManifestResourceNames", ctypes.c_void_p),
        ("GetManifestResourceInfo", ctypes.c_void_p),
        ("get_Location", ctypes.c_void_p),
        ("get_Evidence", ctypes.c_void_p),
        ("GetCustomAttributes", ctypes.c_void_p),
        ("GetCustomAttributes_2", ctypes.c_void_p),
        ("IsDefined", ctypes.c_void_p),
        ("GetObjectData", ctypes.c_void_p),
        ("add_ModuleResolve", ctypes.c_void_p),
        ("remove_ModuleResolve", ctypes.c_void_p),
        ("GetType_4", ctypes.c_void_p),
        ("GetSatelliteAssembly", ctypes.c_void_p),
        ("GetSatelliteAssembly_2", ctypes.c_void_p),
        ("LoadModule", ctypes.c_void_p),
        ("LoadModule_2", ctypes.c_void_p),
        ("CreateInstance", ctypes.c_void_p),
        ("CreateInstance_2", ctypes.c_void_p),
        ("CreateInstance_3", ctypes.c_void_p),
        ("GetLoadedModules", ctypes.c_void_p),
        ("GetLoadedModules_2", ctypes.c_void_p),
        ("GetModules", ctypes.c_void_p),
        ("GetModules_2", ctypes.c_void_p),
        ("GetModule", ctypes.c_void_p),
        ("GetReferencedAssemblies", ctypes.c_void_p),
        ("get_GlobalAssemblyCache", ctypes.c_void_p)
    ]

class Assembly(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(AssemblyVtbl))
    ]

class IEnumUnknownVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("Next", ctypes.c_void_p),
        ("Skip", ctypes.c_void_p),
        ("Reset", ctypes.c_void_p),
        ("Clone", ctypes.c_void_p),
    ]

# Define the IEnumUnknown structure
class IEnumUnknown(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(IEnumUnknownVtbl)),
    ]

class ICLRMetaHostVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("GetRuntime", ctypes.WINFUNCTYPE(
                                    ctypes.c_ulong,
                                    ctypes.c_void_p,	
                                    ctypes.c_wchar_p, 
                                    ctypes.POINTER(GUID), 
                                    ctypes.c_void_p
        )),
        ("GetVersionFromFile", ctypes.c_void_p),
        ("EnumerateInstalledRuntimes", ctypes.WINFUNCTYPE(
                                                    ctypes.c_long, 
                                                    ctypes.c_void_p,
                                                    ctypes.POINTER(IEnumUnknown))
        ),
        ("EnumerateLoadedRuntimes", ctypes.c_void_p),
        ("RequestRuntimeLoadedNotification", ctypes.c_void_p),
        ("QueryLegacyV2RuntimeBinding", ctypes.c_void_p),
        ("ExitProcess", ctypes.c_void_p)
    ]

# Define the ICLRMetaHost structure
class ICLRMetaHost(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(ICLRMetaHostVtbl))
    ]

class ICLRRuntimeInfoVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("GetVersionString", ctypes.c_void_p),
        ("GetRuntimeDirectory", ctypes.c_void_p),
        ("IsLoaded", ctypes.c_void_p),
        ("LoadErrorString", ctypes.c_void_p),
        ("LoadLibrary", ctypes.c_void_p),
        ("GetProcAddress", ctypes.c_void_p),
        ("GetInterface", ctypes.WINFUNCTYPE(
                                    ctypes.c_ulong,
                                    ctypes.c_void_p,
                                    ctypes.POINTER(GUID),
                                    ctypes.POINTER(GUID),
                                    ctypes.c_void_p
        )),
        ("IsLoadable", ctypes.WINFUNCTYPE(
                                    ctypes.c_ulong,
                                    ctypes.c_void_p,
                                    ctypes.POINTER(ctypes.c_bool)
        )),
        ("SetDefaultStartupFlags", ctypes.c_void_p),
        ("GetDefaultStartupFlags", ctypes.c_void_p),
        ("BindAsLegacyV2Runtime", ctypes.c_void_p),
        ("IsStarted", ctypes.c_void_p)
    ]

# Define the ICLRRuntimeInfo structure
class ICLRRuntimeInfo(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(ICLRRuntimeInfoVtbl))
    ]

class AppDomainVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("GetTypeInfoCount", ctypes.c_void_p),
        ("GetTypeInfo", ctypes.c_void_p),
        ("GetIDsOfNames", ctypes.c_void_p),
        ("Invoke", ctypes.c_void_p),
        ("get_ToString", ctypes.c_void_p),
        ("Equals", ctypes.c_void_p),
        ("GetHashCode", ctypes.c_void_p),
        ("GetType", ctypes.c_void_p),
        ("InitializeLifetimeService", ctypes.c_void_p),
        ("GetLifetimeService", ctypes.c_void_p),
        ("get_Evidence", ctypes.c_void_p),
        ("add_DomainUnload", ctypes.c_void_p),
        ("remove_DomainUnload", ctypes.c_void_p),
        ("add_AssemblyLoad", ctypes.c_void_p),
        ("remove_AssemblyLoad", ctypes.c_void_p),
        ("add_ProcessExit", ctypes.c_void_p),
        ("remove_ProcessExit", ctypes.c_void_p),
        ("add_TypeResolve", ctypes.c_void_p),
        ("remove_TypeResolve", ctypes.c_void_p),
        ("add_ResourceResolve", ctypes.c_void_p),
        ("remove_ResourceResolve", ctypes.c_void_p),
        ("add_AssemblyResolve", ctypes.c_void_p),
        ("remove_AssemblyResolve", ctypes.c_void_p),
        ("add_UnhandledException", ctypes.c_void_p),
        ("remove_UnhandledException", ctypes.c_void_p),
        ("DefineDynamicAssembly", ctypes.c_void_p),
        ("DefineDynamicAssembly_2", ctypes.c_void_p),
        ("DefineDynamicAssembly_3", ctypes.c_void_p),
        ("DefineDynamicAssembly_4", ctypes.c_void_p),
        ("DefineDynamicAssembly_5", ctypes.c_void_p),
        ("DefineDynamicAssembly_6", ctypes.c_void_p),
        ("DefineDynamicAssembly_7", ctypes.c_void_p),
        ("DefineDynamicAssembly_8", ctypes.c_void_p),
        ("DefineDynamicAssembly_9", ctypes.c_void_p),
        ("CreateInstance", ctypes.c_void_p),
        ("CreateInstanceFrom", ctypes.c_void_p),
        ("CreateInstance_2", ctypes.c_void_p),
        ("CreateInstanceFrom_2", ctypes.c_void_p),
        ("CreateInstance_3", ctypes.c_void_p),
        ("CreateInstanceFrom_3", ctypes.c_void_p),
        ("Load", ctypes.c_void_p),
        ("Load_2", ctypes.c_void_p),
        ("Load_3", ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p, ctypes.POINTER(SafeArray), ctypes.c_void_p)),
        ("Load_4", ctypes.c_void_p),
        ("Load_5", ctypes.c_void_p),
        ("Load_6", ctypes.c_void_p),
        ("Load_7", ctypes.c_void_p),
        ("ExecuteAssembly", ctypes.c_void_p),
        ("ExecuteAssembly_2", ctypes.c_void_p),
        ("ExecuteAssembly_3", ctypes.c_void_p),
        ("get_FriendlyName", ctypes.c_void_p),
        ("get_BaseDirectory", ctypes.c_void_p),
        ("get_RelativeSearchPath", ctypes.c_void_p),
        ("get_ShadowCopyFiles", ctypes.c_void_p),
        ("GetAssemblies", ctypes.c_void_p),
        ("AppendPrivatePath", ctypes.c_void_p),
        ("ClearPrivatePath", ctypes.c_void_p),
        ("SetShadowCopyPath", ctypes.c_void_p),
        ("ClearShadowCopyPath", ctypes.c_void_p),
        ("SetCachePath", ctypes.c_void_p),
        ("SetData", ctypes.c_void_p),
        ("GetData", ctypes.c_void_p),
        ("SetAppDomainPolicy", ctypes.c_void_p),
        ("SetThreadPrincipal", ctypes.c_void_p),
        ("SetPrincipalPolicy", ctypes.c_void_p),
        ("DoCallBack", ctypes.c_void_p),
        ("get_DynamicDirectory", ctypes.c_void_p)
    ]

# Define the AppDomain structure
class AppDomain(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(AppDomainVtbl))
    ]


class IUnknownVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.WINFUNCTYPE(
                                ctypes.c_ulong, ctypes.c_void_p, ctypes.POINTER(GUID), ctypes.c_void_p)
        ),
        ("AddRef", ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)),
        ("Release", ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p))
    ]

# Define the IUnknown structure
class IUnknown(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(IUnknownVtbl))
    ]

class ICLRRuntimeHostVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("CreateLogicalThreadState", ctypes.c_void_p),
        ("DeleteLogicalThreadState", ctypes.c_void_p),
        ("SwitchInLogicThreadState", ctypes.c_void_p),
        ("SwitchOutLogicalThreadState", ctypes.c_void_p),
        ("LocksHeldByLogicalThreadState", ctypes.c_void_p),
        ("MapFile", ctypes.c_void_p),
        ("GetConfiguration", ctypes.c_void_p),
        ("Start", ctypes.WINFUNCTYPE(
                                ctypes.c_ulong,
                                ctypes.c_void_p
        )),
        ("Stop", ctypes.c_void_p),
        ("CreateDomain", ctypes.c_void_p),
        ("GetDefaultDomain", ctypes.WINFUNCTYPE(
                                ctypes.c_ulong, ctypes.c_void_p, ctypes.c_void_p
        )),
        ("EnumDomains", ctypes.c_void_p),
        ("NextDomain", ctypes.c_void_p),
        ("CloseEnum", ctypes.c_void_p),
        ("CreateDomainEx", ctypes.c_void_p),
        ("CreateDomainSetup", ctypes.c_void_p),
        ("CreateEvidence", ctypes.c_void_p),
        ("UnloadDomain", ctypes.c_void_p),
        ("CurrentDomain", ctypes.c_void_p),
    ]

# Define the ICLRRuntimeHost structure
class ICLRRuntimeHost(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(ICLRRuntimeHostVtbl))
    ]

# Define the VARIANT structure
class VARIANT(ctypes.Structure):
    class U(ctypes.Union):
        _fields_ = [("VT_BOOL", ctypes.c_short),
                    ("VT_BSTR", BSTR),
                    ("VT_DISPATCH", ctypes.c_void_p), #ctypes.POINTER(IDispatch)),
                    ("VT_I1", ctypes.c_char),
                    ("VT_I2", ctypes.c_short),
                    ("VT_I4", ctypes.c_long),
                    ("VT_INT", ctypes.c_int),
                    ("VT_R4", ctypes.c_float),
                    ("VT_R8", ctypes.c_double),
                    ("VT_SCODE", ctypes.c_ulong),
                    ("VT_UI1", ctypes.c_byte),
                    ("VT_UI2", ctypes.c_ushort),
                    ("VT_UI4", ctypes.c_ulong),
                    ("VT_UINT", ctypes.c_uint),
                    ("VT_UNKNOWN", ctypes.POINTER(IUnknown)),
                    # faked fields, only for our convenience:
                    ("wstrVal", ctypes.c_wchar_p),
                    ("voidp", ctypes.c_void_p),
                    ("parray", ctypes.POINTER(SafeArray))
                    ]
    _fields_ = [("vt", ctypes.c_ushort),
                ("wReserved1", ctypes.c_ushort),
                ("wReserved2", ctypes.c_ushort),
                ("wReserved3", ctypes.c_ushort),
                ("_", U)]
    def _set_value(self, value):
        typ = type(value)
        if typ is int:
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_INT
            self._.VT_INT = value
        elif typ is ctypes.POINTER(SafeArray):
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_ARRAY | VT_UI1
            self._.parray = value
        elif typ is str:
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_BSTR
            pValue = ctypes.c_wchar_p(str(value))
            bstrValue = SysAllocString(pValue)
            self._.VT_BSTR = bstrValue
        elif typ is bytes:
            value = value.decode()
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_BSTR
            self._.voidp = SysAllocString(value)
        elif value is None:
            OleAut32.VariantClear(ctypes.byref(self))
        elif typ is bool:
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_BOOL
            self._.VT_BOOL = value and -1 or 0
        elif typ is ctypes.POINTER(IUnknown):
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_UNKNOWN
            self._.VT_UNKNOWN = value
            if value:
                value.AddRef()
        elif typ is ctypes.POINTER(VARIANT):
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_VARIANT
            self._.VT_VARIANT = value
            if value:
                value.AddRef()
        elif typ is ctypes.POINTER(IDispatch):
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_DISPATCH
            self._.VT_DISPATCH = value
            if value:
                value.AddRef()
        elif hasattr(value, "QueryInterface") and issubclass(typ._type_, IDispatch):
            p = ctypes.POINTER(IDispatch)()
            if value:
                value.QueryInterface(ctypes.byref(IDispatch._iid_), byref(p))
                p.AddRef()
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_DISPATCH
            self._.VT_DISPATCH = p
        elif hasattr(value, "QueryInterface") and issubclass(typ._type_, IUnknown):
            p = ctypes.POINTER(IUnknown)()
            if value:
                value.QueryInterface(ctypes.byref(IUnknown._iid_), byref(p))
                p.AddRef()
            OleAut32.VariantClear(ctypes.byref(self))
            self.vt = VT_UNKNOWN
            self._.VT_UNKNOWN = p
        else:
            raise TypeError("don't know how to store %r in a VARIANT" % value)
    def _get_value(self):
        if self.vt == VT_EMPTY:
            return None
        elif self.vt == VT_ARRAY | VT_UI1:
            return self._.parray
        elif self.vt == VT_I1:
            return self._.VT_I1
        elif self.vt == VT_I2:
            return self._.VT_I2
        elif self.vt == VT_I4:
            return self._.VT_I4
        elif self.vt == VT_UI1:
            return self._.VT_UI1
        elif self.vt == VT_UI2:
            return self._.VT_UI2
        elif self.vt == VT_UI4:
            return self._.VT_UI4
        elif self.vt == VT_INT:
            return self._.VT_INT
        elif self.vt == VT_UINT:
            return self._.VT_UINT
        elif self.vt == VT_R4:
            return self._.VT_R4
        elif self.vt == VT_R8:
            return self._.VT_R8
        elif self.vt == VT_BSTR:
            return self._.wstrVal
        elif self.vt == VT_UNKNOWN:
            result = self._.VT_UNKNOWN
            if result:
                result.AddRef()
            return result
        elif self.vt == VT_DISPATCH:
            result = self._.VT_DISPATCH
            if result:
                result.AddRef()
            return result
        elif self.vt == VT_BOOL:
            return bool(self._.VT_BOOL)
        elif self.vt & VT_BYREF:
            var = VARIANT.from_address(self._.voidp)
            return var
        elif self.vt == VT_ERROR:
            return ("Error", self._.VT_SCODE)
        elif self.vt == VT_NULL:
            return None
        else:
            raise TypeError("don't know how to convert typecode %d" % self.vt)
        # not yet done:
        # VT_ARRAY
        # VT_CY
        # VT_DATE
    value = property(_get_value, _set_value)
    def __repr__(self):
        return "<VARIANT 0x%X at %x>" % (self.vt, id(self))
##    def __del__(self, _clear = OleAut32.VariantClear):
##        _clear(ctypes.byref(self))
    def optional(cls):
        var = VARIANT()
        var.vt = VT_ERROR
        var._.VT_SCODE = 0x80020004
        return var
    optional = classmethod(optional)


VARIANT.U._fields_.append(("VT_VARIANT", ctypes.POINTER(VARIANT)))


class _TypeVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("GetTypeInfoCount", ctypes.c_void_p),
        ("GetTypeInfo", ctypes.c_void_p),
        ("GetIDsOfNames", ctypes.c_void_p),
        ("Invoke", ctypes.c_void_p),
        ("get_ToString", ctypes.c_void_p),
        ("Equals", ctypes.c_void_p),
        ("GetHashCode", ctypes.c_void_p),
        ("GetType", ctypes.c_void_p),
        ("get_MemberType", ctypes.c_void_p),
        ("get_name", ctypes.c_void_p),
        ("get_DeclaringType", ctypes.c_void_p),
        ("get_ReflectedType", ctypes.c_void_p),
        ("GetCustomAttributes", ctypes.c_void_p),
        ("GetCustomAttributes_2", ctypes.c_void_p),
        ("IsDefined", ctypes.c_void_p),
        ("get_Guid", ctypes.c_void_p),
        ("get_Module", ctypes.c_void_p),
        ("get_Assembly", ctypes.c_void_p),
        ("get_TypeHandle", ctypes.c_void_p),
        ("get_FullName", ctypes.c_void_p),
        ("get_Namespace", ctypes.c_void_p),
        ("get_AssemblyQualifiedName", ctypes.c_void_p),
        ("GetArrayRank", ctypes.c_void_p),
        ("get_BaseType", ctypes.c_void_p),
        ("GetConstructors", ctypes.c_void_p),
        ("GetInterface", ctypes.c_void_p),
        ("GetInterfaces", ctypes.c_void_p),
        ("FindInterfaces", ctypes.c_void_p),
        ("GetEvent", ctypes.c_void_p),
        ("GetEvents", ctypes.c_void_p),
        ("GetEvents_2", ctypes.c_void_p),
        ("GetNestedTypes", ctypes.c_void_p),
        ("GetNestedType", ctypes.c_void_p),
        ("GetMember", ctypes.c_void_p),
        ("GetDefaultMembers", ctypes.c_void_p),
        ("FindMembers", ctypes.c_void_p),
        ("GetElementType", ctypes.c_void_p),
        ("IsSubclassOf", ctypes.c_void_p),
        ("IsInstanceOfType", ctypes.c_void_p),
        ("IsAssignableFrom", ctypes.c_void_p),
        ("GetInterfaceMap", ctypes.c_void_p),
        ("GetMethod", ctypes.c_void_p),
        ("GetMethod_2", ctypes.c_void_p),
        ("GetMethods", ctypes.c_void_p),
        ("GetField", ctypes.c_void_p),
        ("GetFields", ctypes.c_void_p),
        ("GetProperty", ctypes.c_void_p),
        ("GetProperty_2", ctypes.c_void_p),
        ("GetProperties", ctypes.c_void_p),
        ("GetMember_2", ctypes.c_void_p),
        ("GetMembers", ctypes.c_void_p),
        ("InvokeMember", ctypes.c_void_p),
        ("get_UnderlyingSystemType", ctypes.c_void_p),
        ("InvokeMember_2", ctypes.c_void_p),
        ("InvokeMember_3", ctypes.WINFUNCTYPE(
                                        ctypes.c_ulong,
                                        ctypes.c_void_p,
                                        BSTR,
                                        ctypes.c_uint,
                                        ctypes.c_void_p,
                                        VARIANT,
                                        ctypes.POINTER(SafeArray),
                                        ctypes.POINTER(VARIANT)
        )),
        ("GetConstructor", ctypes.c_void_p),
        ("GetConstructor_2", ctypes.c_void_p),
        ("GetConstructor_3", ctypes.c_void_p),
        ("GetConstructors_2", ctypes.c_void_p),
        ("get_TypeInitializer", ctypes.c_void_p),
        ("GetMethod_3", ctypes.c_void_p),
        ("GetMethod_4", ctypes.c_void_p),
        ("GetMethod_5", ctypes.c_void_p),
        ("GetMethod_6", ctypes.c_void_p),
        ("GetMethods_2", ctypes.c_void_p),
        ("GetField_2", ctypes.c_void_p),
        ("GetFields_2", ctypes.c_void_p),
        ("GetInterface_2", ctypes.c_void_p),
        ("GetEvent_2", ctypes.c_void_p),
        ("GetProperty_3", ctypes.c_void_p),
        ("GetProperty_4", ctypes.c_void_p),
        ("GetProperty_5", ctypes.c_void_p),
        ("GetProperty_6", ctypes.c_void_p),
        ("GetProperty_7", ctypes.c_void_p),
        ("GetProperties_2", ctypes.c_void_p),
        ("GetNestedTypes_2", ctypes.c_void_p),
        ("GetNestedType_2", ctypes.c_void_p),
        ("GetMember_3", ctypes.c_void_p),
        ("GetMembers_2", ctypes.c_void_p),
        ("get_Attributes", ctypes.c_void_p),
        ("get_IsNotPublic", ctypes.c_void_p),
        ("get_IsPublic", ctypes.c_void_p),
        ("get_IsNestedPublic", ctypes.c_void_p),
        ("get_IsNestedPrivate", ctypes.c_void_p),
        ("get_IsNestedFamily", ctypes.c_void_p),
        ("get_IsNestedAssembly", ctypes.c_void_p),
        ("get_IsNestedFamANDAssem", ctypes.c_void_p),
        ("get_IsNestedFamORAssem", ctypes.c_void_p),
        ("get_IsAutoLayout", ctypes.c_void_p),
        ("get_IsLayoutSequential", ctypes.c_void_p),
        ("get_IsExplicitLayout", ctypes.c_void_p),
        ("get_IsClass", ctypes.c_void_p),
        ("get_IsInterface", ctypes.c_void_p),
        ("get_IsValueType", ctypes.c_void_p),
        ("get_IsAbstract", ctypes.c_void_p),
        ("get_IsSealed", ctypes.c_void_p),
        ("get_IsEnum", ctypes.c_void_p),
        ("get_IsSpecialName", ctypes.c_void_p),
        ("get_IsImport", ctypes.c_void_p),
        ("get_IsSerializable", ctypes.c_void_p),
        ("get_IsAnsiClass", ctypes.c_void_p),
        ("get_IsUnicodeClass", ctypes.c_void_p),
        ("get_IsAutoClass", ctypes.c_void_p),
        ("get_IsArray", ctypes.c_void_p),
        ("get_IsByRef", ctypes.c_void_p),
        ("get_IsPointer", ctypes.c_void_p),
        ("get_IsPrimitive", ctypes.c_void_p),
        ("get_IsCOMObject", ctypes.c_void_p),
        ("get_HasElementType", ctypes.c_void_p),
        ("get_IsContextful", ctypes.c_void_p),
        ("get_IsMarshalByRef", ctypes.c_void_p),
        ("Equals_2", ctypes.c_void_p)
    ]

# Define the _Type structure
class Type(ctypes.Structure):
    _fields_ = [
        ("vtbl", ctypes.POINTER(_TypeVtbl))
    ]

# COM GUIDs
CLSID_CLRMetaHost = GUID(0x9280188d, 0xe8e, 0x4867, (ctypes.c_ubyte * 8)(0xb3, 0xc, 0x7f, 0xa8, 0x38, 0x84, 0xe8, 0xde))
CLSID_CorRuntimeHost = GUID(0xcb2f6723, 0xab3a, 0x11d2, (ctypes.c_ubyte * 8)(0x9c, 0x40, 0x00, 0xc0, 0x4f, 0xa3, 0x0a, 0x3e))
IID_ICLRMetaHost = GUID(0xD332DB9E, 0xB9B3, 0x4125, (ctypes.c_ubyte * 8)(0x82, 0x07, 0xA1, 0x48, 0x84, 0xF5, 0x32, 0x16))
IID_ICLRRuntimeInfo = GUID(0xBD39D1D2, 0xBA2F, 0x486a, (ctypes.c_ubyte * 8)(0x89, 0xB0, 0xB4, 0xB0, 0xCB, 0x46, 0x68, 0x91))
IID_ICorRuntimeHost = GUID(0xcb2f6722, 0xab3a, 0x11d2, (ctypes.c_ubyte * 8)(0x9c, 0x40, 0x00, 0xc0, 0x4f, 0xa3, 0x0a, 0x3e))
IID_AppDomain = GUID(0x05F696DC, 0x2B29, 0x3663, (ctypes.c_ubyte * 8)(0xAD, 0x8B, 0xC4,0x38, 0x9C, 0xF2, 0xA7, 0x13))

# Windows API Functions
ntdll = ctypes.WinDLL("ntdll.dll")
RtlMoveMemory = ntdll.RtlMoveMemory
RtlMoveMemory.argtypes = (
        ctypes.c_void_p,
        ctypes.c_void_p,
        ctypes.c_size_t
)
RtlMoveMemory.restype = ctypes.c_void_p
OleAut32 = ctypes.WinDLL("OleAut32.dll")
SafeArrayCreate = OleAut32.SafeArrayCreate
SafeArrayCreate.restype = ctypes.POINTER(SafeArray)
SafeArrayCreate.argtypes = [
    ctypes.c_ushort,
    ctypes.c_uint,
    ctypes.POINTER(SafeArrayBound * 1)
]
SafeArrayCreateVector = OleAut32.SafeArrayCreateVector
SafeArrayCreateVector.restype = ctypes.POINTER(SafeArray)
SafeArrayCreateVector.argtypes = [
    ctypes.c_ushort,
    ctypes.c_long,
    ctypes.c_ulong
]
SafeArrayDestroy = OleAut32.SafeArrayDestroy
SafeArrayDestroy.argtypes = (
    ctypes.POINTER(SafeArray),
)
SafeArrayPutElement = OleAut32.SafeArrayPutElement
SafeArrayPutElement.restype = ctypes.c_ulong
SafeArrayPutElement.argtypes = (
    ctypes.POINTER(SafeArray),
    ctypes.POINTER(ctypes.c_long),
    ctypes.c_void_p
)
SysAllocString = OleAut32.SysAllocString
SysAllocString.argtypes = [
    ctypes.c_wchar_p
]
SysAllocString.restype = BSTR
mscoree = ctypes.WinDLL("mscoree.dll")
CLRCreateInstance = mscoree.CLRCreateInstance
CLRCreateInstance.argtypes = (
        ctypes.POINTER(GUID),
        ctypes.POINTER(GUID),
        ctypes.c_void_p
)
