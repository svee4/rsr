
using System;
using System.Data;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using Windows.Win32.Media.Audio;
using Windows.Win32.System.Com;
using Windows.Win32.System.Com.StructuredStorage;
using Windows.Win32.UI.Shell.PropertiesSystem;

// https://github.com/frgnca/AudioDeviceCmdlets/tree/master/SOURCE

var devices = GetDevices();

Console.WriteLine("Devices found:");
PrintDevices(devices);

List<int> deviceIndexes = [];

while (true)
{
    Console.WriteLine($"Choose devices by index. Currently selected: {(deviceIndexes.Count is 0 ? "(none)" : string.Join(", ", deviceIndexes))}");

    var line = Console.ReadLine();

    if (string.IsNullOrEmpty(line))
    {
        break;
    }

    if (!int.TryParse(line, out var deviceIndex))
    {
        Console.WriteLine("Could not parse device index, input a number");
        continue;
    }

    if (deviceIndexes.Contains(deviceIndex))
    {
        deviceIndexes.Remove(deviceIndex);
    } 
    else
    {
        deviceIndexes.Add(deviceIndex);
    }
}

if (deviceIndexes.Count is 0) 
{
    Console.WriteLine("No devices selected, quitting...");
    Console.ReadLine();
    return;
}

List<Device> selectedDevices = devices.Select((d, i) => (d, i)).Where(t => deviceIndexes.Contains(t.i + 1)).Select(t => t.d).ToList();

Console.WriteLine("Devices selected: ");
foreach (var device in selectedDevices)
{
    Console.WriteLine($"\t{device.Name}");
}

Console.WriteLine("---------------");

Console.WriteLine("""
Options:
    1. Set sample rate to 48000
    2. Set sample rate to 44100
    3. Quit without changes
""");

if (!int.TryParse(Console.ReadLine(), out var choice) || choice is not (1 or 2))
{
    Console.WriteLine("Quitting without changes...");
    Console.ReadLine();
    return;
}

var newSampleRate = choice is 1 ? 48000 : 44100;

foreach (var device in selectedDevices)
{
    var format = device.Format;
    if (format.SubFormat != new Guid("00000001-0000-0010-8000-00aa00389b71"))
    {
        throw new InvalidOperationException($"Device '{device.Name}' wave sub format unexpected - only PCM can be handled");
    }

    format.Format.nSamplesPerSec = (uint)newSampleRate;
    format.Format.nAvgBytesPerSec = (uint)(newSampleRate * format.Format.nBlockAlign);

    var prop = new PROPVARIANT();
    prop.Anonymous.Anonymous.vt = Windows.Win32.System.Variant.VARENUM.VT_BLOB;

    unsafe
    {
        var formatStructSize = Marshal.SizeOf(format);
        var formatStructPtr = NativeMemory.Alloc((nuint)formatStructSize);

        Marshal.StructureToPtr(format, (nint)formatStructPtr, false);

        prop.Anonymous.Anonymous.Anonymous.blob = new BLOB
        {
            cbSize = (uint)formatStructSize,
            pBlobData = (byte*)formatStructPtr
        };

        var deviceFormatKey = Keys.DeviceFormat;
        device.PropertyStore.SetValue(&deviceFormatKey, prop);
        device.PropertyStore.Commit();

        NativeMemory.Free(formatStructPtr);
    }
}


Console.WriteLine("Changes saved. New state: ");
PrintDevices(GetDevices());
Console.WriteLine("Quitting...");
Console.ReadLine();


unsafe static List<Device> GetDevices()
{

    var deviceEnuemrator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
    deviceEnuemrator.EnumAudioEndpoints(EDataFlow.eCapture, DEVICE_STATE.DEVICE_STATE_ACTIVE, out var deviceCollection);
    deviceCollection.GetCount(out var numDevices);

    List<Device> devices = new((int)numDevices);

    for (int i = 0; i < numDevices; i++)
    {
        deviceCollection.Item((uint)i, out var device);
        device.OpenPropertyStore(STGM.STGM_READWRITE, out var store);

        var deviceFriendlyNameKey = Keys.DeviceFriendlyName;
        store.GetValue(&deviceFriendlyNameKey, out var nameProp);
        var name = Marshal.PtrToStringUni((nint)nameProp.Anonymous.Anonymous.Anonymous.pszVal.Value);
        
        if (name is null)
        {
            Console.WriteLine("Null name, skipping");
            continue;
        }

        var deviceFormatKey = Keys.DeviceFormat;
        store.GetValue(&deviceFormatKey, out var formatProp);

        if (formatProp.Anonymous.Anonymous.vt != Windows.Win32.System.Variant.VARENUM.VT_BLOB)
        {
            throw new InvalidOperationException($"PKEY_AudioEngine_DeviceFormat PROPVARIANT type is not VT_BLOB. Actual type: '{formatProp.Anonymous.Anonymous.vt}'");
        }

        WAVEFORMATEXTENSIBLE format = Marshal.PtrToStructure<WAVEFORMATEXTENSIBLE>((nint)formatProp.Anonymous.Anonymous.Anonymous.blob.pBlobData);

        devices.Add(new Device()
        {
            Name = name,
            SampleRate = format.Format.nSamplesPerSec,
            Channels = format.Format.nChannels,
            SampleSize = format.Format.wBitsPerSample,
            Format = format,
            PropertyStore = store
        });
    }

    return devices;
}

static void PrintDevices(List<Device> devices)
{
    var longestName = devices.Select(d => d.Name.Length).Max();

    foreach (var (device, index) in devices.Select((d, i) => (d, i)))
    {
        Console.WriteLine($"{index + 1}: {device.Name}: {new string(' ', longestName - device.Name.Length)}{device.SampleRate} ({device.Channels} channel, {device.SampleSize}bit)");
    }
}

static class Keys
{
    public static PROPERTYKEY DeviceFormat => new() { fmtid = new Guid(0xf19f064d, 0x82c, 0x4e27, 0xbc, 0x73, 0x68, 0x82, 0xa1, 0xbb, 0x8e, 0x4c), pid = 0 };
    public static PROPERTYKEY DeviceFriendlyName => new() { fmtid = new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 14 };
}

[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
public class MMDeviceEnumerator;

internal class Device
{
    public required string Name { get; init; }
    public required uint SampleRate { get; init; }
    public required int Channels { get; init; }
    public required int SampleSize { get; init; }
    public required WAVEFORMATEXTENSIBLE Format { get; init; }
    public required IPropertyStore PropertyStore { get; init; }
}