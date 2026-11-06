// Research helper based on Microsoft's ApplicationLoopback API sample.
// Captures only the explicit target process tree, never a microphone or endpoint-wide mix.
using System;using System.IO;using System.Runtime.InteropServices;using System.Threading;
namespace OfficeCliMediaControls {
 [ComImport,Guid("72A22D78-CDE4-431D-B8CC-843A71199B6D"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface IAsyncOperation {[PreserveSig]int GetActivateResult(out int result,[MarshalAs(UnmanagedType.IUnknown)]out object client);}
 [ComVisible(true),Guid("41D949AB-9862-444A-80F6-C261334DA5EB"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface ICompletion {[PreserveSig]int ActivateCompleted(IAsyncOperation operation);}
 [ComVisible(true),Guid("94EA2B94-E9CC-49E0-C0FF-EE64CA8F5B90"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]public interface IAgile {}
 [StructLayout(LayoutKind.Sequential,Pack=2)]public struct WaveFormat {public ushort tag,channels;public uint rate,bytesPerSecond;public ushort blockAlign,bits,extra;}
 [ComImport,Guid("1CB9AD4C-DBFA-4C32-B178-C2F568A703B2"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface IAudioClient {
  [PreserveSig]int Initialize(int share,uint flags,long duration,long periodicity,ref WaveFormat format,IntPtr session);
  [PreserveSig]int GetBufferSize(out uint frames);[PreserveSig]int GetStreamLatency(out long value);[PreserveSig]int GetCurrentPadding(out uint frames);
  [PreserveSig]int IsFormatSupported(int share,IntPtr format,out IntPtr closest);[PreserveSig]int GetMixFormat(out IntPtr format);[PreserveSig]int GetDevicePeriod(out long standard,out long minimum);
  [PreserveSig]int Start();[PreserveSig]int Stop();[PreserveSig]int Reset();[PreserveSig]int SetEventHandle(IntPtr handle);
  [PreserveSig]int GetService(ref Guid id,[MarshalAs(UnmanagedType.IUnknown)]out object service);
 }
 [ComImport,Guid("C8ADBD64-E71E-48A0-A4DE-185C395CD317"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface ICaptureClient {
  [PreserveSig]int GetBuffer(out IntPtr data,out uint frames,out uint flags,out ulong devicePosition,out ulong qpcPosition);
  [PreserveSig]int ReleaseBuffer(uint frames);[PreserveSig]int GetNextPacketSize(out uint frames);
 }
 [StructLayout(LayoutKind.Explicit,Size=24)]public struct PropVariant {[FieldOffset(0)]public ushort type;[FieldOffset(8)]public uint size;[FieldOffset(16)]public IntPtr data;}
 [ComVisible(true),ClassInterface(ClassInterfaceType.None)]public sealed class Completion:ICompletion,IAgile {
  public readonly ManualResetEvent Done=new ManualResetEvent(false);public object Client;public int Result;
  public int ActivateCompleted(IAsyncOperation op){try{int inner;Result=op.GetActivateResult(out inner,out Client);if(Result>=0)Result=inner;}catch(Exception e){Result=Marshal.GetHRForException(e);}finally{Done.Set();}return 0;}
 }
 public sealed class ProcessLoopback:IDisposable {
  [DllImport("Mmdevapi.dll",CharSet=CharSet.Unicode,ExactSpelling=true)]static extern int ActivateAudioInterfaceAsync(string path,ref Guid iid,ref PropVariant parameters,ICompletion completion,out IAsyncOperation operation);
  [DllImport("ole32.dll")]static extern int CoInitializeEx(IntPtr reserved,uint apartment);[DllImport("ole32.dll")]static extern void CoUninitialize();
  readonly uint pid;readonly string path;readonly ManualResetEvent ready=new ManualResetEvent(false);Thread worker;volatile bool stopping;Exception error;long bytes;public int Packets,Discontinuities,TimestampErrors;public ulong FirstQpc,LastQpc;public int TargetPid {get{return (int)pid;}}public long AudioBytes {get{return bytes;}}public string Failure {get{return error==null?null:error.ToString();}}
  public ProcessLoopback(int processId,string output){if(processId<=0)throw new ArgumentOutOfRangeException("processId");if(File.Exists(output))throw new IOException("Fresh audio path required");pid=(uint)processId;path=output;}
  static void Check(int hr){if(hr<0)Marshal.ThrowExceptionForHR(hr);}
  public void Start(){worker=new Thread(Capture);worker.IsBackground=true;worker.SetApartmentState(ApartmentState.MTA);worker.Start();if(!ready.WaitOne(10000))throw new TimeoutException("Process loopback start timeout");if(error!=null)throw new InvalidOperationException("Process loopback activation failed",error);}
  public void Stop(){stopping=true;if(worker!=null&&!worker.Join(5000))throw new TimeoutException("Process loopback stop timeout");if(error!=null)throw new InvalidOperationException("Process loopback failed",error);}
  void Capture(){
   int co=CoInitializeEx(IntPtr.Zero,0);IAudioClient client=null;ICaptureClient capture=null;IAsyncOperation operation=null;IntPtr activation=IntPtr.Zero;bool started=false;
   try{
    Check(co);if(IntPtr.Size!=8)throw new PlatformNotSupportedException("x64 research helper");
    activation=Marshal.AllocHGlobal(12);Marshal.WriteInt32(activation,0,1);Marshal.WriteInt32(activation,4,(int)pid);Marshal.WriteInt32(activation,8,0);
    PropVariant pv=new PropVariant{type=65,size=12,data=activation};Guid audioId=typeof(IAudioClient).GUID;Completion done=new Completion();Check(ActivateAudioInterfaceAsync("VAD\\Process_Loopback",ref audioId,ref pv,done,out operation));
    if(!done.Done.WaitOne(7000))throw new TimeoutException("Native audio activation callback timeout");Check(done.Result);client=(IAudioClient)done.Client;
    WaveFormat format=new WaveFormat{tag=1,channels=2,rate=44100,bits=16,blockAlign=4,bytesPerSecond=176400,extra=0};
    // Same shared PCM conversion and process-loopback flags as Microsoft's sample; polling replaces its MF event queue.
    Check(client.Initialize(0,0x80020000,0,0,ref format,IntPtr.Zero));Guid captureId=typeof(ICaptureClient).GUID;object service;Check(client.GetService(ref captureId,out service));capture=(ICaptureClient)service;
    using(FileStream file=new FileStream(path,FileMode.CreateNew,FileAccess.Write,FileShare.Read))using(BinaryWriter writer=new BinaryWriter(file)){
     writer.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"));writer.Write(0);writer.Write(System.Text.Encoding.ASCII.GetBytes("WAVEfmt "));writer.Write(16);writer.Write((ushort)1);writer.Write((ushort)2);writer.Write(44100);writer.Write(176400);writer.Write((ushort)4);writer.Write((ushort)16);writer.Write(System.Text.Encoding.ASCII.GetBytes("data"));writer.Write(0);
     Check(client.Start());started=true;ready.Set();
     while(!stopping){uint packet;Check(capture.GetNextPacketSize(out packet));while(packet>0){IntPtr data;uint frames,flags;ulong position,qpc;Check(capture.GetBuffer(out data,out frames,out flags,out position,out qpc));try{if(Packets==0)FirstQpc=qpc;else if((flags&1)!=0)Discontinuities++;LastQpc=qpc;Packets++;if((flags&4)!=0)TimestampErrors++;byte[] buffer=new byte[checked((int)frames*4)];if((flags&2)==0)Marshal.Copy(data,buffer,0,buffer.Length);writer.Write(buffer);bytes+=buffer.Length;}finally{Check(capture.ReleaseBuffer(frames));}Check(capture.GetNextPacketSize(out packet));}Thread.Sleep(5);}
     Check(client.Stop());started=false;writer.Seek(4,SeekOrigin.Begin);writer.Write(checked((int)bytes+36));writer.Seek(40,SeekOrigin.Begin);writer.Write(checked((int)bytes));writer.Flush();
    }
   }catch(Exception e){error=e;ready.Set();}
   finally{if(started&&client!=null)try{client.Stop();}catch{};if(capture!=null)Marshal.ReleaseComObject(capture);if(client!=null)Marshal.ReleaseComObject(client);if(operation!=null)Marshal.ReleaseComObject(operation);if(activation!=IntPtr.Zero)Marshal.FreeHGlobal(activation);if(co>=0)CoUninitialize();}
  }
  public void Dispose(){if(worker!=null&&worker.IsAlive)Stop();ready.Dispose();}
 }
}
