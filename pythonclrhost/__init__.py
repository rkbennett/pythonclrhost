import ctypes
from .windef import *

class clrhost(object):
    def __init__(self):
        self.pAssembly = ctypes.c_void_p()
        self.Class = None
        self.ConfigFile = None

    def get_installed_runtimes(self, metaHost):
        pInstalledRuntimes = IEnumUnknown()
        HRESULT = metaHost.contents.vtbl.contents.EnumerateInstalledRuntimes(
            metaHost,
            ctypes.byref(pInstalledRuntimes)
        )
        if HRESULT:
            raise OSError(f"Failed to enumerate runtimes with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")

    def load(self, buf: bytes=None, rtime: str=None):
        sz = ctypes.c_uint(1)
        pUnk = ctypes.c_void_p()
        bLoadable = ctypes.c_bool()
        clr_ptr = ctypes.c_uint64()
        pIUnknown = ctypes.c_void_p()
        pAppDomain = ctypes.c_void_p()
        assemblyPtr = ctypes.c_void_p()
        pRuntimeInfo = ctypes.c_void_p(None)
        charArray = ctypes.create_string_buffer(buf)
        bounds = (SafeArrayBound * 1)()
        bounds[0].cElements = len(charArray)
        bounds[0].lLbound = 0
        HRESULT = CLRCreateInstance(
            ctypes.byref(CLSID_CLRMetaHost), 
            ctypes.byref(IID_ICLRMetaHost), 
            ctypes.byref(clr_ptr)
        )
        if HRESULT:
            raise OSError(f"Failed to create CLR instance with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        pMetaHost = ctypes.cast(clr_ptr.value, ctypes.POINTER(ICLRMetaHost))
        if not rtime:
            self.get_installed_runtimes(pMetaHost)
        else:
            runtime = ctypes.c_wchar_p(rtime)
        HRESULT = pMetaHost.contents.vtbl.contents.GetRuntime(
            pMetaHost,
            runtime, 
            ctypes.byref(IID_ICLRRuntimeInfo),
            ctypes.byref(pRuntimeInfo)
        )
        if HRESULT:
            raise OSError(f"Failed to get runtime with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        RuntimeInfo = ctypes.cast(pRuntimeInfo, ctypes.POINTER(ICLRRuntimeInfo))
        HRESULT = RuntimeInfo.contents.vtbl.contents.IsLoadable(
            RuntimeInfo,
            ctypes.byref(bLoadable)
        )
        if HRESULT:
            raise OSError(f"Failed to verify if runtime is loadable with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        if not bLoadable.value:
            raise OSError("Runtime is not loadable")
        HRESULT = RuntimeInfo.contents.vtbl.contents.GetInterface(
            RuntimeInfo,
            ctypes.byref(CLSID_CorRuntimeHost),
            ctypes.byref(IID_ICorRuntimeHost),
            ctypes.byref(pUnk)
        )
        if HRESULT:
            raise OSError(f"Failed to get interface with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        Unk = ctypes.cast(pUnk, ctypes.POINTER(ICLRRuntimeHost))
        HRESULT = Unk.contents.vtbl.contents.Start(
            Unk
        )
        if HRESULT:
            raise OSError(f"Failed to start runtime with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = Unk.contents.vtbl.contents.GetDefaultDomain(
            Unk,
            ctypes.byref(pIUnknown)
        )
        if HRESULT:
            raise OSError(f"Failed to get default domain with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        iUnknown = ctypes.cast(pIUnknown, ctypes.POINTER(IUnknown))
        HRESULT = iUnknown.contents.vtbl.contents.QueryInterface(
            iUnknown,
            ctypes.byref(IID_AppDomain),
            ctypes.byref(pAppDomain)
        )
        if HRESULT:
            raise OSError(f"Failed to query interface with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        appDomain = ctypes.cast(pAppDomain, ctypes.POINTER(AppDomain))
        safe_array = SafeArrayCreate(VT_UI1,sz,ctypes.byref(bounds))
        if not safe_array:
            raise OSError(f"Failed to create safearray with error: {ctypes.get_last_error()}")
        RtlMoveMemory(
            safe_array.contents.pvData, 
            charArray, 
            len(charArray)
        )
        HRESULT = appDomain.contents.vtbl.contents.Load_3(
            appDomain,
            safe_array,
            ctypes.byref(assemblyPtr)
        )
        self.pAssembly = ctypes.cast(assemblyPtr, ctypes.POINTER(Assembly))

    def manage_domain(self, Class: str, Method: str):
        vtEmpty = VARIANT()
        self.Class = Class
        pType = ctypes.c_void_p()
        cls = ctypes.c_wchar_p(self.Class)
        nullPtr = ctypes.c_void_p(None)
        vtPSEntryPointReturnVal = VARIANT()
        safe_array = SafeArrayCreateVector(VT_VARIANT,0,0)
        bstrCls = SysAllocString(
            cls
        )
        HRESULT = self.pAssembly.contents.vtbl.contents.GetType_2(
            self.pAssembly,
            bstrCls,
            ctypes.byref(pType)
        )
        if HRESULT:
            raise OSError(f"Failed to get type with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        dotnetMethod = ctypes.c_wchar_p(Method)
        bstrDotnetMethod = SysAllocString(
            dotnetMethod
        )
        typeInterface = ctypes.cast(pType, ctypes.POINTER(Type))
        HRESULT = typeInterface.contents.vtbl.contents.InvokeMember_3(
            typeInterface,
            bstrDotnetMethod,
            ctypes.c_uint(BindingFlags.BindingFlags_InvokeMethod | BindingFlags.BindingFlags_Static | BindingFlags.BindingFlags_Public),
            nullPtr,
            vtEmpty,
            safe_array,
            ctypes.byref(vtPSEntryPointReturnVal)
        )
        if HRESULT:
            raise OSError(f"Failed to invoke member with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayDestroy(safe_array)
        if HRESULT:
            raise OSError(f"Failed to clean up safearray with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")

    def create_appdomain(self, Domain: str, ConfigFile: str=None):
        vtEmpty = VARIANT()
        pType = ctypes.c_void_p()
        self.ConfigFile = ConfigFile
        cls = ctypes.c_wchar_p(self.Class)
        nullPtr = ctypes.c_void_p(None)
        domain = ctypes.c_wchar_p(Domain)
        vtPSEntryPointReturnVal = VARIANT()
        bstrCls = SysAllocString(
            cls
        )
        staticMethodName = ctypes.c_wchar_p("CreateAppDomain")
        bstrStaticMethodName = SysAllocString(
            staticMethodName
        )
        bstrDomain = SysAllocString(
            domain
        )
        configFile = ctypes.c_wchar_p(self.ConfigFile)
        bstrConfigFile = SysAllocString(
            configFile
        )
        vtStringArg = VARIANT()
        vtStringArg._set_value(bstrDomain)
        vtStringArg2 = VARIANT()
        if bstrConfigFile:
            vtStringArg2._set_value(bstrConfigFile)
        HRESULT= self.pAssembly.contents.vtbl.contents.GetType_2(
            self.pAssembly,
            bstrCls,
            ctypes.byref(pType)
        )
        if HRESULT:
            raise OSError(f"Failed to get type with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        typeInterface = ctypes.cast(pType, ctypes.POINTER(Type))
        psaStaticMethodArgs = SafeArrayCreateVector(VT_VARIANT,0,2)
        if not psaStaticMethodArgs:
            raise OSError(f"Failed to create safearray vector with error: {ctypes.get_last_error()}")
        index0 = ctypes.c_long(0)
        index1 = ctypes.c_long(1)
        HRESULT = SafeArrayPutElement(
            psaStaticMethodArgs,
            ctypes.byref(index0),
            ctypes.byref(vtStringArg)
        )
        if HRESULT:
            raise OSError(f"Failed to put safearray element with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayPutElement(
            psaStaticMethodArgs,
            ctypes.byref(index1),
            ctypes.byref(vtStringArg2)
        )
        if HRESULT:
            raise OSError(f"Failed to put safearray element with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = typeInterface.contents.vtbl.contents.InvokeMember_3(
            typeInterface,
            bstrStaticMethodName,
            ctypes.c_uint(BindingFlags.BindingFlags_InvokeMethod | BindingFlags.BindingFlags_Static | BindingFlags.BindingFlags_Public),
            nullPtr,
            vtEmpty,
            psaStaticMethodArgs,
            ctypes.byref(vtPSEntryPointReturnVal)
        )
        if HRESULT:
            raise OSError(f"Failed to invoke member with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayDestroy(psaStaticMethodArgs)
        if HRESULT:
            raise OSError(f"Failed to clean up safearray with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        if vtPSEntryPointReturnVal._get_value() == 0:
            raise OSError(f"Failed to create appdomain")
        vtPSEntryPointReturnVal.vt = VT_UINT
        return vtPSEntryPointReturnVal._get_value()
    
    def get_function(self, buff: bytes, Domain: int=0):
        vtEmpty = VARIANT()
        sz = ctypes.c_uint(1)
        pType = ctypes.c_void_p()
        nullPtr = ctypes.c_void_p(None)
        cls = ctypes.c_wchar_p(self.Class)
        vtPSEntryPointReturnVal = VARIANT()
        asm = ctypes.create_string_buffer(buff)
        typeName = ctypes.c_wchar_p("Python.Runtime.Loader")
        bstrTypeName = SysAllocString(
            typeName
        )
        function = ctypes.c_wchar_p("Initialize")
        bstrFunction = SysAllocString(
            function
        )
        vtTypeArg = VARIANT()
        vtTypeArg._set_value(bstrTypeName)
        vtFuncArg = VARIANT()
        vtFuncArg._set_value(bstrFunction)
        cls = ctypes.c_wchar_p("ClrLoader.ClrLoader")
        bstrCls = SysAllocString(
            cls
        )
        staticMethodName = "GetFunction"
        bstrStaticMethodName = SysAllocString(
            staticMethodName
        )
        domainArg = VARIANT()
        domainArg._set_value(int(Domain))
        asmContent = (SafeArrayBound * 1)()
        asmContent[0].cElements = len(asm)
        asmContent[0].lLbound = 0
        asmSafeArr = SafeArrayCreate(VT_UI1,sz,ctypes.byref(asmContent))
        if not asmSafeArr:
            raise OSError(f"Failed to create safearray with error: {ctypes.get_last_error()}")
        vtAsmArg = VARIANT()
        vtAsmArg._set_value(asmSafeArr)
        RtlMoveMemory(asmSafeArr.contents.pvData, asm, len(asm))
        HRESULT = self.pAssembly.contents.vtbl.contents.GetType_2(
            self.pAssembly,
            bstrCls,
            ctypes.byref(pType)
        )
        if HRESULT:
            raise OSError(f"Failed to get type with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        typeInterface = ctypes.cast(pType, ctypes.POINTER(Type))
        psaStaticMethodArgs = SafeArrayCreateVector(ctypes.c_uint16(VT_VARIANT),0,4)
        if not psaStaticMethodArgs:
            raise OSError(f"Failed to create safearray vector with error: {ctypes.get_last_error()}")
        index0 = ctypes.c_long(0)
        index1 = ctypes.c_long(1)
        index2 = ctypes.c_long(2)
        index3 = ctypes.c_long(3)
        HRESULT = SafeArrayPutElement(
            psaStaticMethodArgs,
            ctypes.byref(index0),
            ctypes.byref(domainArg)
        )
        if HRESULT:
            raise OSError(f"Failed to put safearray element with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayPutElement(
            psaStaticMethodArgs,
            ctypes.byref(index1),
            ctypes.byref(vtAsmArg)
        )
        if HRESULT:
            raise OSError(f"Failed to put safearray element with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayPutElement(
            psaStaticMethodArgs,
            ctypes.byref(index2),
            ctypes.byref(vtTypeArg)
        )
        if HRESULT:
            raise OSError(f"Failed to put safearray element with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayPutElement(
            psaStaticMethodArgs,
            ctypes.byref(index3),
            ctypes.byref(vtFuncArg)
        )
        if HRESULT:
            raise OSError(f"Failed to put safearray element with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = typeInterface.contents.vtbl.contents.InvokeMember_3(
            typeInterface,
            bstrStaticMethodName,
            ctypes.c_uint(BindingFlags.BindingFlags_InvokeMethod | BindingFlags.BindingFlags_Static | BindingFlags.BindingFlags_Public),
            nullPtr,
            vtEmpty,
            psaStaticMethodArgs,
            ctypes.byref(vtPSEntryPointReturnVal)
        )
        if HRESULT:
            raise OSError(f"Failed to invoke member with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayDestroy(psaStaticMethodArgs)
        if HRESULT:
            raise OSError(f"Failed to clean up safearray with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        vtPSEntryPointReturnVal.vt = VT_UINT
        return vtPSEntryPointReturnVal._get_value()
    
    def close_appdomain(self, Domain: int):
        vtEmpty = VARIANT()
        pType = ctypes.c_void_p()
        cls = ctypes.c_wchar_p(self.Class)
        nullPtr = ctypes.c_void_p(None)
        vtPSEntryPointReturnVal = VARIANT()
        bstrCls = SysAllocString(
            cls
        )
        HRESULT = self.pAssembly.contents.vtbl.contents.GetType_2(
            self.pAssembly,
            bstrCls,
            ctypes.byref(pType)
        )
        if HRESULT:
            raise OSError(f"Failed to get type with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        psaStaticMethodArgs = SafeArrayCreateVector(VT_VARIANT,0,1)
        if not psaStaticMethodArgs:
            raise OSError(f"Failed to create safearray vector with error: {ctypes.get_last_error()}")
        index0 = ctypes.c_long(0)
        typeInterface = ctypes.cast(pType, ctypes.POINTER(Type))
        domainArg = VARIANT()
        domainArg._set_value(int(Domain))
        staticMethodName = ctypes.c_wchar_p("CloseAppDomain")
        bstrStaticMethodName = SysAllocString(
            staticMethodName
        )
        HRESULT = SafeArrayPutElement(
            psaStaticMethodArgs,
            ctypes.byref(index0),
            ctypes.byref(domainArg)
        )
        if HRESULT:
            raise OSError(f"Failed to put safearray element with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = typeInterface.contents.vtbl.contents.InvokeMember_3(
            typeInterface,
            bstrStaticMethodName,
            ctypes.c_uint(BindingFlags.BindingFlags_InvokeMethod | BindingFlags.BindingFlags_Static | BindingFlags.BindingFlags_Public),
            nullPtr,
            vtEmpty,
            psaStaticMethodArgs,
            ctypes.byref(vtPSEntryPointReturnVal)
        )
        if HRESULT:
            raise OSError(f"Failed to invoke member with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")
        HRESULT = SafeArrayDestroy(psaStaticMethodArgs)
        if HRESULT:
            raise OSError(f"Failed to clean up safearray with error: {win_errors[str(HRESULT)]} ({hex(HRESULT)})")

    def pyclr_initialize(self):
        self.manage_domain("ClrLoader.ClrLoader", "Initialize")

    def pyclr_finalize(self):
        self.manage_domain("ClrLoader.ClrLoader", "Close")
