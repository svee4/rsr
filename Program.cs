
using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

// https://github.com/frgnca/AudioDeviceCmdlets/tree/master/SOURCE

PropertyKey PKEY_AudioEngine_DeviceFormat = new PropertyKey { fmtid = new Guid(0xf19f064d, 0x82c, 0x4e27, 0xbc, 0x73, 0x68, 0x82, 0xa1, 0xbb, 0x8e, 0x4c), pid = 0 };
PropertyKey PKEY_DeviceInterface_FriendlyName = new PropertyKey { fmtid = new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 14 };

Console.WriteLine("Hello world");


var deviceEnuemrator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
IMMDeviceCollection deviceCollection;

Marshal.ThrowExceptionForHR(deviceEnuemrator.EnumAudioEndpoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE, out deviceCollection));
Marshal.ThrowExceptionForHR(deviceCollection.GetCount(out var numDevices));

for (int i = 0; i < numDevices; i++)
{
    Marshal.ThrowExceptionForHR(deviceCollection.Item((uint)i, out var device));
    Marshal.ThrowExceptionForHR(device.OpenPropertyStore(EStgmAccess.STGM_READWRITE, out var store));

    Marshal.ThrowExceptionForHR(store.GetValue(ref PKEY_DeviceInterface_FriendlyName, out var propVariant));
    Console.WriteLine(propVariant.Value);

    Marshal.ThrowExceptionForHR(store.GetValue(ref PKEY_AudioEngine_DeviceFormat, out var propVariant2));

    Console.WriteLine(propVariant2.Value is null ? "NO" : "YES");

    if (propVariant2.Value is byte[] arr)
    {
        var s = new WAVEFORMATEXTENSIBLE();

        int size = Marshal.SizeOf(s);
        IntPtr ptr = IntPtr.Zero;
        try
        {
            ptr = Marshal.AllocHGlobal(size);

            Marshal.Copy(arr, 0, ptr, size);

            s = (WAVEFORMATEXTENSIBLE)Marshal.PtrToStructure(ptr, s.GetType())!;
            Console.WriteLine("samples: " + s.Format.nSamplesPerSec);
        }
        finally
        {
            Marshal.FreeHGlobal(ptr);
        }
    }

    //Marshal.ThrowExceptionForHR(store.GetCount(out var storeCount));
    //for (int j = 0; j < storeCount; j++)
    //{
    //    Marshal.ThrowExceptionForHR(store.GetAt(j, out var pkey));
    //    if (pkey.fmtid == PKEY_AudioEngine_DeviceFormat)
    //    {
    //        Console.WriteLine("THINGY FOUND: " + pkey.pid);
    //    }
    //    Marshal.ThrowExceptionForHR(store.GetValue(ref pkey, out var value));
    //    Console.WriteLine($"    {pkey.fmtid}: {value.Value}");
    //}
}


[Guid("D666063F-1587-4E43-81F1-B948E807363F")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMMDevice
{
    int Activate(ref Guid iid, CLSCTX dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
    int OpenPropertyStore(EStgmAccess stgmAccess, out IPropertyStore propertyStore);
    int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
    int GetState(out EDeviceState pdwState);
}

[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMMDeviceEnumerator
{
    int EnumAudioEndpoints(EDataFlow dataFlow, EDeviceState StateMask, out IMMDeviceCollection device);
    //int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
    //int GetDevice(string pwstrId, out IMMDevice ppDevice);
    //int RegisterEndpointNotificationCallback(IntPtr pClient);
    //int UnregisterEndpointNotificationCallback(IntPtr pClient);
}

[Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPropertyStore
{
    [PreserveSig]
    int GetCount(out Int32 count);
    [PreserveSig]
    int GetAt(int iProp, out PropertyKey pkey);
    [PreserveSig]
    int GetValue(ref PropertyKey key, out PropVariant pv);
    [PreserveSig]
    int SetValue(ref PropertyKey key, ref PropVariant propvar);
    [PreserveSig]
    int Commit();
};

[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
public class MMDeviceEnumerator;

[StructLayout(LayoutKind.Explicit)]
public struct PropVariant
{
    [FieldOffset(0)] short vt;
    [FieldOffset(2)] short wReserved1;
    [FieldOffset(4)] short wReserved2;
    [FieldOffset(6)] short wReserved3;
    [FieldOffset(8)] sbyte cVal;
    [FieldOffset(8)] byte bVal;
    [FieldOffset(8)] short iVal;
    [FieldOffset(8)] ushort uiVal;
    [FieldOffset(8)] int lVal;
    [FieldOffset(8)] uint ulVal;
    [FieldOffset(8)] long hVal;
    [FieldOffset(8)] ulong uhVal;
    [FieldOffset(8)] float fltVal;
    [FieldOffset(8)] double dblVal;
    [FieldOffset(8)] Blob blobVal;
    [FieldOffset(8)] DateTime date;
    [FieldOffset(8)] bool boolVal;
    [FieldOffset(8)] int scode;
    [FieldOffset(8)] System.Runtime.InteropServices.ComTypes.FILETIME filetime;
    [FieldOffset(8)] IntPtr everything_else;

    //I'm sure there is a more efficient way to do this but this works ..for now..
    public byte[] GetBlob()
    {
        byte[] Result = new byte[blobVal.Length];
        for (int i = 0; i < blobVal.Length; i++)
        {
            Result[i] = Marshal.ReadByte((IntPtr)((long)(blobVal.Data) + i));
        }
        return Result;
    }

    public object? Value
    {
        get
        {
            VarEnum ve = (VarEnum)vt;
            switch (ve)
            {
                case VarEnum.VT_I1:
                    return bVal;
                case VarEnum.VT_I2:
                    return iVal;
                case VarEnum.VT_I4:
                    return lVal;
                case VarEnum.VT_I8:
                    return hVal;
                case VarEnum.VT_INT:
                    return iVal;
                case VarEnum.VT_UI4:
                    return ulVal;
                case VarEnum.VT_LPWSTR:
                    return Marshal.PtrToStringUni(everything_else);
                case VarEnum.VT_BLOB:
                    return GetBlob();
            }

            return null;
        }
    }

}

public struct Blob
{
    public int Length;
    public IntPtr Data;
}

public struct PropertyKey
{
    public Guid fmtid;
    public int pid;
};

public enum EDataFlow
{
    eRender = 0,
    eCapture = 1,
    eAll = 2,
    EDataFlow_enum_count = 3
}

[Flags]
public enum EDeviceState : uint
{
    DEVICE_STATE_ACTIVE = 0x00000001,
    DEVICE_STATE_UNPLUGGED = 0x00000002,
    DEVICE_STATE_NOTPRESENT = 0x00000004,
    DEVICE_STATEMASK_ALL = 0x00000007
}

[Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMMDeviceCollection
{
    int GetCount(out uint pcDevices);
    int Item(uint nDevice, out IMMDevice Device);
}

[Flags]
public enum CLSCTX : uint
{
    INPROC_SERVER = 0x1,
    INPROC_HANDLER = 0x2,
    LOCAL_SERVER = 0x4,
    INPROC_SERVER16 = 0x8,
    REMOTE_SERVER = 0x10,
    INPROC_HANDLER16 = 0x20,
    RESERVED1 = 0x40,
    RESERVED2 = 0x80,
    RESERVED3 = 0x100,
    RESERVED4 = 0x200,
    NO_CODE_DOWNLOAD = 0x400,
    RESERVED5 = 0x800,
    NO_CUSTOM_MARSHAL = 0x1000,
    ENABLE_CODE_DOWNLOAD = 0x2000,
    NO_FAILURE_LOG = 0x4000,
    DISABLE_AAA = 0x8000,
    ENABLE_AAA = 0x10000,
    FROM_DEFAULT_CONTEXT = 0x20000,
    INPROC = INPROC_SERVER | INPROC_HANDLER,
    SERVER = INPROC_SERVER | LOCAL_SERVER | REMOTE_SERVER,
    ALL = SERVER | INPROC_HANDLER
}

public enum EStgmAccess
{
    STGM_READ = 0x00000000,
    STGM_WRITE = 0x00000001,
    STGM_READWRITE = 0x00000002
}

[StructLayout(LayoutKind.Sequential, Pack = 1)]
public struct WAVEFORMATEX
{
    public ushort wFormatTag;
    public ushort nChannels;
    public uint nSamplesPerSec;
    public uint nAvgBytesPerSec;
    public ushort nBlockAlign;
    public ushort wBitsPerSample;
    public ushort cbSize;
}

[StructLayout(LayoutKind.Sequential, Pack = 1)]
public struct WAVEFORMATEXTENSIBLE
{
    public WAVEFORMATEX Format;
    public ushort union;
    public uint dwChannelMask;
    public Guid SubFormat;
}